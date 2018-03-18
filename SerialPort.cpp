/*
**  FILENAME            SerialPort.cpp
**
**  PURPOSE             This class can read, write and watch one serial port.
**                      It sends messages to its owner when something happends on the port
**                      The class creates a thread for reading and writing so the main
**                      program is not blocked.
**
**  CREATION DATE       15-09-1997
**  LAST MODIFICATION   18-03-2018
**
**  AUTHOR              Remon Spekreijse
*/

#include "stdafx.h"
#include "SerialPort.h"
#include <assert.h>

#pragma warning(disable:4996)

CSerialPort::CSerialPort()
{
    m_hComm = INVALID_HANDLE_VALUE;
    m_szWriteBuffer = NULL;
    m_bThreadAlive = FALSE;
    m_bUserRequestClose = FALSE;
    m_nWriteSize = 0;
    m_Thread = NULL;
    InitializeCriticalSection( &m_csCommunicationSync );
}

CSerialPort::~CSerialPort()
{
    Close();
    DeleteCriticalSection( &m_csCommunicationSync );
}

BOOL CSerialPort::Open( HWND    pPortOwner,      // the owner (CWnd) of the port (receives message)
                        UINT    port,            // portnumber (0..SERIAL_PORT_MAX)
                        UINT    baud,            // baudrate
                        BYTE    parity,          // parity
                        BYTE    databits,        // databits
                        BYTE    stopbits,        // stopbits
                        DWORD   dwCommEvents,    // EV_RXCHAR, EV_CTS etc
                        UINT    nBufferSize,     // size to the writebuffer
                        DWORD   ReadIntervalTimeout,
                        DWORD   ReadTotalTimeoutMultiplier,
                        DWORD   ReadTotalTimeoutConstant,
                        DWORD   WriteTotalTimeoutMultiplier,
                        DWORD   WriteTotalTimeoutConstant )

{
    BOOL ret = TRUE;
    char szPort[MAX_PATH];
    Close();
    EnterCriticalSection( &m_csCommunicationSync );
    assert( port <= SERIAL_PORT_MAX );
    assert( pPortOwner != NULL );
    // save the owner
    m_pOwner = pPortOwner;
    // Allocate memory
    m_szWriteBuffer = new char[nBufferSize];

    if ( m_szWriteBuffer == NULL )
    {
        ret = FALSE;
        goto done;
    }

    m_nPortNr = port;
    m_nWriteBufferSize = nBufferSize;
    m_dwCommEvents = dwCommEvents | (EV_RXCHAR | EV_TXEMPTY);
    // prepare port strings
    sprintf( szPort, _T( "\\\\.\\%s%d" ), SERIAL_DEVICE_PREFIX, (signed int)port );
    // get a handle to the port
    m_hComm = CreateFile( szPort,                       // communication port string (COMX)
                          GENERIC_READ | GENERIC_WRITE, // read/write types
                          0,                            // comm devices must be opened with exclusive access
                          NULL,                         // no security attributes
                          OPEN_EXISTING,                // comm devices must use OPEN_EXISTING
                          FILE_FLAG_NO_BUFFERING | FILE_FLAG_WRITE_THROUGH,
                          0 );                          // template must be 0 for comm devices

    if ( m_hComm == INVALID_HANDLE_VALUE )
    {
        ret = FALSE;
        goto done;
    }

    // set the timeout values
    m_CommTimeouts.ReadIntervalTimeout       = ReadIntervalTimeout;
    m_CommTimeouts.ReadTotalTimeoutMultiplier  = ReadTotalTimeoutMultiplier;
    m_CommTimeouts.ReadTotalTimeoutConstant = ReadTotalTimeoutConstant;
    m_CommTimeouts.WriteTotalTimeoutMultiplier = WriteTotalTimeoutMultiplier;
    m_CommTimeouts.WriteTotalTimeoutConstant   = WriteTotalTimeoutConstant;

    // configure
    if ( SetCommTimeouts( m_hComm, &m_CommTimeouts ) )
    {
        if ( SetCommMask( m_hComm, dwCommEvents ) )
        {
            if ( GetCommState( m_hComm, &m_dcb ) )
            {
                m_dcb.BaudRate = baud;
                m_dcb.Parity   = parity;
                m_dcb.ByteSize = databits;
                m_dcb.StopBits = stopbits;
                m_dcb.fOutxCtsFlow = FALSE;
                m_dcb.fRtsControl = RTS_CONTROL_DISABLE;
                m_dcb.fOutxDsrFlow = FALSE;
                m_dcb.fDtrControl = DTR_CONTROL_DISABLE;
                m_dcb.fBinary = TRUE;
                m_dcb.fDsrSensitivity = FALSE;
                m_dcb.fTXContinueOnXoff = FALSE;
                m_dcb.fOutX = FALSE;
                m_dcb.fInX = FALSE;
                m_dcb.fErrorChar = FALSE;
                m_dcb.fNull = FALSE;
                m_dcb.fAbortOnError = FALSE;

                if ( SetCommState( m_hComm, &m_dcb ) == 0 )
                {
                    ProcessErrorMessage( "SetCommState()" );
                    ret = FALSE;
                    goto done;
                }
            }
            else
            {
                ProcessErrorMessage( "GetCommState()" );
                ret = FALSE;
                goto done;
            }
        }
        else
        {
            ProcessErrorMessage( "SetCommMask()" );
            ret = FALSE;
            goto done;
        }
    }
    else
    {
        ProcessErrorMessage( "SetCommTimeouts()" );
        ret = FALSE;
        goto done;
    }

    // set the SetupComm parameter into device control.
    if ( !SetupComm( m_hComm, m_nWriteBufferSize, m_nWriteBufferSize ) )
    {
        ProcessErrorMessage( "SetupComm()" );
        ret = FALSE;
        goto done;
    }

    // flush the port
    if ( !PurgeComm( m_hComm, PURGE_RXCLEAR | PURGE_TXCLEAR | PURGE_RXABORT | PURGE_TXABORT ) )
    {
        ProcessErrorMessage( "PurgeComm()" );
        ret = FALSE;
        goto done;
    }

    m_bThreadAlive = TRUE;
    m_bUserRequestClose = FALSE;
    assert( m_Thread == NULL );
    m_Thread = ::CreateThread(NULL, 0, CommThread, this, 0, NULL);

    if (m_Thread == NULL)
    {
        ProcessErrorMessage( "CreateThread()" );
        ret = FALSE;
        m_bThreadAlive = FALSE;
        goto done;
    }

done:
    LeaveCriticalSection( &m_csCommunicationSync );

    if ( !ret )
    {
        Close();
    }

    return ret;
}

DWORD WINAPI CSerialPort::CommThread( LPVOID pParam )
{
    CSerialPort *pPort = ( CSerialPort * )pParam;

    while ( pPort->m_bThreadAlive )
    {
        if ( pPort->m_bUserRequestClose )
        {
            break;
        }

        if (pPort->m_nWriteSize > 0)
        {
            UINT BytesSent = WriteChar(pPort);

            if (EOF != BytesSent)
            {
                ::PostMessage(pPort->m_pOwner, SERIAL_PORT_MESSAGE, ( WPARAM )EV_TXEMPTY, ( LPARAM )BytesSent);
            }
            else
            {
                break;
            }
        }

        ReceiveChar(pPort);
    }

    pPort->m_bThreadAlive = FALSE;
    ::ExitThread(0);
    //return 0;
}

void CSerialPort::ProcessErrorMessage( char *ErrorText )
{
    char *Temp;
    char *lpMsgBuf = NULL;
    DWORD dwError = GetLastError();

    if ( FormatMessage(
                FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM,
                NULL,
                dwError,
                MAKELANGID( LANG_NEUTRAL, SUBLANG_DEFAULT ),
                ( LPTSTR )&lpMsgBuf,
                0,
                NULL
            ) > 0 )
    {
        if (lpMsgBuf != NULL)
        {
            Temp = ( char * )LocalAlloc( LMEM_ZEROINIT, ( lstrlen( ( LPCTSTR )lpMsgBuf ) + INT16_MAX ) );

            if ( Temp != NULL )
            {
                sprintf( Temp, _T( "ERROR:\"%s\" failed with the following error:\n\ndwError=%d\n%s\nPort:%s%d\n" ), ( char * )ErrorText, (signed int)dwError, lpMsgBuf, SERIAL_DEVICE_PREFIX, (signed int)m_nPortNr );
                AfxMessageBox(Temp, MB_OK | MB_ICONERROR | MB_SYSTEMMODAL | MB_TOPMOST);
                LocalFree( Temp );
            }

            LocalFree( lpMsgBuf );
        }

        assert(dwError == ERROR_SUCCESS);
    }
}

UINT CSerialPort::WriteChar( CSerialPort *pPort )
{
    BOOL bResult;
    DWORD Sent = 0;
    EnterCriticalSection( &pPort->m_csCommunicationSync );
    bResult = WriteFile( pPort->m_hComm,
                         pPort->m_szWriteBuffer,
                         pPort->m_nWriteSize,
                         &Sent,
                         NULL);
    LeaveCriticalSection( &pPort->m_csCommunicationSync );

    if ( !bResult )
    {
        pPort->ProcessErrorMessage("WriteFile()");
        pPort->m_nWriteSize = 0;
        return (UINT)EOF; //break;
    }
    else
    {
        assert ( Sent == (DWORD)pPort->m_nWriteSize );
        pPort->m_nWriteSize = 0;
        return (UINT)Sent;
    }
}

BOOL CSerialPort::ReceiveChar( CSerialPort *pPort )
{
    BOOL  bResult = TRUE;
    DWORD BytesRead = 0;
    unsigned char RXBuff;

    while ( pPort->m_bThreadAlive )
    {
        if ( pPort->m_bUserRequestClose )
        {
            break;
        }

        EnterCriticalSection( &pPort->m_csCommunicationSync );
        bResult = ReadFile( pPort->m_hComm,      // Handle to COMM port
                            &RXBuff,             // RX Buffer Pointer
                            sizeof( RXBuff ),    // Read one byte
                            &BytesRead,          // Stores number of bytes read
                            NULL);
        LeaveCriticalSection( &pPort->m_csCommunicationSync );

        if (bResult && (BytesRead > 0) && (BytesRead <= sizeof(RXBuff)))
        {
            ::PostMessage( pPort->m_pOwner, SERIAL_PORT_MESSAGE, ( WPARAM )  EV_RXCHAR, ( LPARAM ) RXBuff );
        }
        else if ((!bResult) && (ERROR_ACCESS_DENIED == GetLastError()))
        {
            pPort->ProcessErrorMessage("ReadFile()");
            return FALSE;
        }
        else
        {
            ::Sleep( MAX_PATH );
            break;
        }
    }

    return TRUE;
}

DCB *CSerialPort::GetDCB()
{
    return &m_dcb;
}

BOOL CSerialPort::SetDCB( DCB *dcb )
{
    BOOL ret = TRUE;
    assert( m_hComm != INVALID_HANDLE_VALUE );
    assert( dcb != NULL );

    do
    {
        ::Sleep( 0 );
    }
    while ( m_nWriteSize );

    EnterCriticalSection( &m_csCommunicationSync );
    m_dcb = *dcb;

    if ( SetCommState( m_hComm, &m_dcb ) == 0 )
    {
        ProcessErrorMessage( "SetCommState()" );
        ret = FALSE;
    }

    LeaveCriticalSection( &m_csCommunicationSync );
    return ret;
}

BOOL CSerialPort::IsOpen()
{
    return m_hComm != INVALID_HANDLE_VALUE;
}

void CSerialPort::Close()
{
    if ( (!m_bUserRequestClose) && (m_bThreadAlive) )
    {
        m_bUserRequestClose = TRUE;

        do
        {
            ::Sleep( 0 );
        }
        while ( m_bThreadAlive );

        m_bUserRequestClose = FALSE;
    }

    EnterCriticalSection( &m_csCommunicationSync );

    if ( m_hComm != INVALID_HANDLE_VALUE )
    {
        CloseHandle( m_hComm );
        m_hComm = INVALID_HANDLE_VALUE;
    }

    if ( m_szWriteBuffer != NULL )
    {
        delete [] m_szWriteBuffer;
        m_szWriteBuffer = NULL;
    }

    if ( m_Thread != NULL )
    {
        CloseHandle( m_Thread );
        m_Thread = NULL;
    }

    LeaveCriticalSection( &m_csCommunicationSync );
}

void CSerialPort::Write( char *Buffer )
{
    int nSize;
    int ttl;
    assert( m_hComm != INVALID_HANDLE_VALUE );
    assert( Buffer != NULL );
    nSize = strlen( Buffer );
    assert( nSize > 0 );
    EnterCriticalSection( &m_csCommunicationSync );
    assert( ( m_nWriteSize + nSize ) < ( int )m_nWriteBufferSize );
    strcpy( m_szWriteBuffer + m_nWriteSize, Buffer );
    m_nWriteSize += nSize;
    ttl = ((1000UL * (m_dcb.ByteSize + m_dcb.StopBits + 1) * m_nWriteSize) / m_dcb.BaudRate) + 1;
    LeaveCriticalSection( &m_csCommunicationSync );

    do
    {
        ::Sleep( 0 );
    }
    while ( m_nWriteSize );

    ::Sleep( ttl );
}

void CSerialPort::Write( void *Buffer, int nSize )
{
    int ttl;
    assert( m_hComm != INVALID_HANDLE_VALUE );
    assert( Buffer != NULL );
    assert( nSize > 0 );
    EnterCriticalSection( &m_csCommunicationSync );
    assert( ( m_nWriteSize + nSize ) < ( int )m_nWriteBufferSize );
    memcpy( m_szWriteBuffer + m_nWriteSize, Buffer, nSize );
    m_nWriteSize += nSize;
    ttl = ((1000UL * (m_dcb.ByteSize + m_dcb.StopBits + 1) * m_nWriteSize) / m_dcb.BaudRate) + 1;
    LeaveCriticalSection( &m_csCommunicationSync );

    do
    {
        ::Sleep( 0 );
    }
    while ( m_nWriteSize );

    ::Sleep( ttl );
}

BOOL CSerialPort::QueryRegistry( HKEY hKey )
{
    TCHAR    achClass[MAX_PATH] = _T( "" );   // buffer for class name
    DWORD    cchClassName = MAX_PATH;         // size of class string
    DWORD    cSubKeys = 0;                    // number of subkeys
    DWORD    cbMaxSubKey;                     // longest subkey size
    DWORD    cchMaxClass;                     // longest class string
    DWORD    cValues = 0;                     // number of values for key
    DWORD    cchMaxValue;                     // longest value name
    DWORD    cbMaxValueData;                  // longest value data
    DWORD    cbSecurityDescriptor;            // size of security descriptor
    FILETIME ftLastWriteTime;                 // last write time
    DWORD    i;
    DWORD    retCode;
    TCHAR    *achValue = NULL;
    DWORD    cchValue;
    BOOL     ret = FALSE;
    achValue = ( TCHAR * )LocalAlloc( LMEM_ZEROINIT, MAX_VALUE_NAME * sizeof( TCHAR ) );

    if ( achValue != NULL )
    {
        retCode = RegQueryInfoKey(
                      hKey,                    // key handle
                      achClass,                // buffer for class name
                      &cchClassName,           // size of class string
                      NULL,                    // reserved
                      &cSubKeys,               // number of subkeys
                      &cbMaxSubKey,            // longest subkey size
                      &cchMaxClass,            // longest class string
                      &cValues,                // number of values for this key
                      &cchMaxValue,            // longest value name
                      &cbMaxValueData,         // longest value data
                      &cbSecurityDescriptor,   // security descriptor
                      &ftLastWriteTime );      // last write time

        if ( retCode == ERROR_SUCCESS )
        {
            for ( i = 0; i < sizeof( m_nComArray ) / sizeof( m_nComArray[0] ); i++ )
            {
                m_nComArray[i] = -1;
            }

            if ( cValues > 0 )
            {
                for ( i = 0; i < cValues; i++ )
                {
                    cchValue = MAX_VALUE_NAME;
                    achValue[0] = '\0';

                    if ( ERROR_SUCCESS == RegEnumValue( hKey, i, achValue, &cchValue, NULL, NULL, NULL, NULL ) )
                    {
                        CString szName( achValue );

                        if ( szName.MakeUpper().Trim().Find( CString( _T( "\\Device\\" ) ).MakeUpper().Trim() ) == 0 )
                        {
                            DWORD    nValueType = 0;
                            BYTE     strDSName[MAX_PATH];
                            DWORD    nBuffLen = sizeof( strDSName );
                            memset( strDSName, 0, sizeof( strDSName ) );

                            if ( ERROR_SUCCESS == RegQueryValueEx( hKey, ( LPCTSTR )achValue, NULL, &nValueType, strDSName, &nBuffLen ) )
                            {
                                int nIndex = 0;

                                while ( nIndex < SERIAL_PORT_MAX && CString( strDSName ).MakeUpper().Trim().Find( CString( SERIAL_DEVICE_PREFIX ).MakeUpper().Trim() ) == 0 )
                                {
                                    if ( -1 == m_nComArray[nIndex] )
                                    {
                                        m_nComArray[nIndex] = atoi( ( char * )( strDSName + strlen( SERIAL_DEVICE_PREFIX ) ) );
                                        break;
                                    }

                                    nIndex++;
                                }
                            }
                        }
                    }
                }

                ret = TRUE;
            }
        }
        else
        {
            OutputDebugString( _T( "Failed to query registry!" ) );
            assert( FALSE );
        }

        LocalFree( achValue );
    }

    return ret;
}

void CSerialPort::EnumSerialPort( CComboBox &m_PortNO )
{
    HKEY hTestKey;
    bool Flag = FALSE;
    int  i = 0;

    if ( ERROR_SUCCESS == RegOpenKeyEx( HKEY_LOCAL_MACHINE, _T( "HARDWARE\\DEVICEMAP\\SERIALCOMM" ), 0, KEY_READ, &hTestKey ) )
    {
        if ( QueryRegistry( hTestKey ) )
        {
            m_PortNO.ResetContent();

            while ( ( i < SERIAL_PORT_MAX ) && ( -1 != m_nComArray[i] ) )
            {
                CString szCom;
                szCom.Format( _T( "%s%d" ), SERIAL_DEVICE_PREFIX, m_nComArray[i] );
                m_PortNO.InsertString( i, szCom.GetBuffer() );
                i++;

                if ( !Flag )
                {
                    Flag = TRUE;
                    m_PortNO.SetCurSel( 0 );
                }
            }
        }

        RegCloseKey( hTestKey );
    }
}
