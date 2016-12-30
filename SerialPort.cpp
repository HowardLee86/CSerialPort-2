/*
**  FILENAME            SerialPort.cpp
**
**  PURPOSE             This class can read, write and watch one serial port.
**                      It sends messages to its owner when something happends on the port
**                      The class creates a thread for reading and writing so the main
**                      program is not blocked.
**
**  CREATION DATE       15-09-1997
**  LAST MODIFICATION   12-30-2016
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
	m_ov.Offset = 0;
	m_ov.OffsetHigh = 0;
	m_ov.hEvent = NULL;
	m_hWriteEvent = NULL;
	m_hShutdownEvent = NULL;
	m_szWriteBuffer = NULL;
	m_bThreadAlive = FALSE;
	m_nWriteSize = 0;
	m_Thread = NULL;
	memset( m_szTitle, 0, sizeof( m_szTitle ) );
}

CSerialPort::~CSerialPort()
{
	if ( IsOpen() )
	{
		Close();
	}
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
	assert( port >= 0 && port <= SERIAL_PORT_MAX );
	assert( pPortOwner != NULL );
	// initialize critical section
	InitializeCriticalSection( &m_csCommunicationSync );
	EnterCriticalSection( &m_csCommunicationSync );
	// save the owner
	m_pOwner = pPortOwner;
	// setlocale(LC_ALL,"");
	GetWindowText( m_pOwner, ( LPSTR )m_szTitle, sizeof( m_szTitle ) );

	// Close port if already opened
	if ( IsOpen() )
	{
		Close();
	}

	// create events
	m_ov.hEvent = CreateEvent( NULL, TRUE, FALSE, NULL );

	if ( m_ov.hEvent == NULL )
	{
		ret = FALSE;
		goto done;
	}

	m_hWriteEvent = CreateEvent( NULL, TRUE, FALSE, NULL );

	if ( m_hWriteEvent == NULL )
	{
		ret = FALSE;
		goto done;
	}

	m_hShutdownEvent = CreateEvent( NULL, TRUE, FALSE, NULL );

	if ( m_hShutdownEvent == NULL )
	{
		ret = FALSE;
		goto done;
	}

	// initialize the event objects
	m_hEventArray[EVENT_SHUTDOWN] = m_hShutdownEvent;
	m_hEventArray[EVENT_READ] = m_ov.hEvent;
	m_hEventArray[EVENT_WRITE] = m_hWriteEvent;
	// Allocate memory
	m_szWriteBuffer = new char[nBufferSize];

	if ( m_szWriteBuffer == NULL )
	{
		ret = FALSE;
		goto done;
	}

	m_nPortNr = port;
	m_nWriteBufferSize = nBufferSize;
	m_dwCommEvents = dwCommEvents;
	// prepare port strings
	sprintf( szPort, _T( "\\\\.\\%s%d" ), SERIAL_DEVICE_PREFIX, port );
	// get a handle to the port
	m_hComm = CreateFile( szPort,                       // communication port string (COMX)
						  GENERIC_READ | GENERIC_WRITE, // read/write types
						  0,                            // comm devices must be opened with exclusive access
						  NULL,                         // no security attributes
						  OPEN_EXISTING,                // comm devices must use OPEN_EXISTING
						  FILE_FLAG_OVERLAPPED,         // Async I/O
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

	assert( m_Thread == NULL );

	if ( !( m_Thread = ::CreateThread ( NULL, 0, CommThread, this, 0, NULL ) ) )
	{
		ProcessErrorMessage( "CreateThread()" );
		ret = FALSE;
		goto done;
	}

done:

	if ( !ret )
	{
		Close();
	}

	LeaveCriticalSection( &m_csCommunicationSync );
	return ret;
}

DWORD WINAPI CSerialPort::CommThread( LPVOID pParam )
{
	DWORD BytesTransfered = 0;
	DWORD Event = 0;
	DWORD CommEvent = 0;
	DWORD dwError = 0;
	COMSTAT comstat;
	BOOL  bResult = TRUE;
	// Cast the void pointer passed to the thread back to
	// a pointer of CSerialPort class
	CSerialPort *pPort = ( CSerialPort * )pParam;
	// Set the status variable in the dialog class to
	// TRUE to indicate the thread is running.
	pPort->m_bThreadAlive = TRUE;

	while ( pPort->m_bThreadAlive )
	{
		// Make a call to WaitCommEvent().  This call will return immediatly
		// because our port was created as an async port (FILE_FLAG_OVERLAPPED
		// and an m_OverlappedStructerlapped structure specified).  This call will cause the
		// m_OverlappedStructerlapped element m_OverlappedStruct.hEvent, which is part of the m_hEventArray to
		// be placed in a non-signeled state if there are no bytes available to be read,
		// or to a signeled state if there are bytes available.  If this event handle
		// is set to the non-signeled state, it will be set to signeled when a
		// character arrives at the port.
		// we do this for each port!
		bResult = WaitCommEvent( pPort->m_hComm, &Event, &pPort->m_ov );

		if ( !bResult )
		{
			// If WaitCommEvent() returns FALSE, process the last error to determin
			// the reason..
			switch ( dwError = GetLastError() )
			{
				case ERROR_IO_PENDING:
				case ERROR_INVALID_PARAMETER:
					{
						// This is a normal return value if there are no bytes
						// to read at the port.
						// Do nothing and continue
						break;
					}

				default:
					{
						// All other error codes indicate a serious error has
						// occured.  Process this error.
						pPort->ProcessErrorMessage( "WaitCommEvent()" );
						break;
					}
			}
		}
		else
		{
			// If WaitCommEvent() returns TRUE, check to be sure there are
			// actually bytes in the buffer to read.
			//
			// If you are reading more than one byte at a time from the buffer
			// (which this program does not do) you will have the situation occur
			// where the first byte to arrive will cause the WaitForMultipleObjects()
			// function to stop waiting.  The WaitForMultipleObjects() function
			// resets the event handle in m_OverlappedStruct.hEvent to the non-signelead state
			// as it returns.
			//
			// If in the time between the reset of this event and the call to
			// ReadFile() more bytes arrive, the m_OverlappedStruct.hEvent handle will be set again
			// to the signeled state. When the call to ReadFile() occurs, it will
			// read all of the bytes from the buffer, and the program will
			// loop back around to WaitCommEvent().
			//
			// At this point you will be in the situation where m_OverlappedStruct.hEvent is set,
			// but there are no bytes available to read.  If you proceed and call
			// ReadFile(), it will return immediatly due to the async port setup, but
			// GetOverlappedResults() will not return until the next character arrives.
			//
			// It is not desirable for the GetOverlappedResults() function to be in
			// this state.  The thread shutdown event (event 0) and the WriteFile()
			// event (Event2) will not work if the thread is blocked by GetOverlappedResults().
			//
			// The solution to this is to check the buffer with a call to ClearCommError().
			// This call will reset the event handle, and if there are no bytes to read
			// we can loop back through WaitCommEvent() again, then proceed.
			// If there are really bytes to read, do nothing and proceed.
			bResult = ClearCommError( pPort->m_hComm, &dwError, &comstat );

			if ( comstat.cbInQue == 0 )
			{
				continue;
			}
		}

		// Main wait function.  This function will normally block the thread
		// until one of nine events occur that require action.
		Event = WaitForMultipleObjects( EVENT_TYPE_MAX,
										pPort->m_hEventArray,
										FALSE,
										INFINITE );

		switch ( Event )
		{
			case EVENT_SHUTDOWN:
				{
					// Shutdown event.  This is event zero so it will be
					// the higest priority and be serviced first.
					pPort->m_bThreadAlive = FALSE;
					::ExitThread( 0 );
					break;
				}

			case EVENT_READ:
				{
					GetCommMask( pPort->m_hComm, &CommEvent );

					if ( CommEvent & EV_RXCHAR )
					{
						ReceiveChar( pPort );
					}

					if ( CommEvent & ( EV_CTS | EV_DSR | EV_RLSD | EV_RXFLAG | EV_BREAK | EV_ERR | EV_RING ) )
					{
						::PostMessage( pPort->m_pOwner, SERIAL_PORT_MESSAGE, ( WPARAM ) 0, ( LPARAM ) ( ( CommEvent & ~EV_RXCHAR ) & ~EV_TXEMPTY ) );
					}

					break;
				}

			case EVENT_WRITE:
				{
					DWORD BytesSent = WriteChar( pPort );
					::PostMessage( pPort->m_pOwner, SERIAL_PORT_MESSAGE, ( WPARAM ) BytesSent, ( LPARAM )  EV_TXEMPTY  );
					break;
				}

			default:
				{
					MessageBox( pPort->m_pOwner, _T( "WaitForMultipleObjects() returned unexpected event:" ), pPort->m_szTitle, MB_ICONERROR );
					break;
				}
		}
	}

	return 0;
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
		Temp = ( char * )LocalAlloc( LMEM_ZEROINIT, ( lstrlen( ( LPCTSTR )lpMsgBuf ) + 1024 ) );

		if ( Temp != NULL )
		{
			sprintf( Temp, _T( "ERROR:\"%s\" failed with the following error:\n\ndwError=%d\n%s\nPort:%s%d\n" ), ( char * )ErrorText, dwError, lpMsgBuf, SERIAL_DEVICE_PREFIX, m_nPortNr );
			MessageBox( m_pOwner, Temp, m_szTitle, MB_ICONERROR );
			LocalFree( Temp );
		}

		if ( lpMsgBuf != NULL )
		{
			LocalFree( lpMsgBuf );
		}
	}
}

DWORD CSerialPort::WriteChar( CSerialPort *pPort )
{
	BOOL bWrite = TRUE;
	BOOL bResult = TRUE;
	DWORD BytesSentA = 0, BytesSentB = 0;
	EnterCriticalSection( &pPort->m_csCommunicationSync );
	ResetEvent( pPort->m_hWriteEvent );
	pPort->m_ov.Offset = 0;
	pPort->m_ov.OffsetHigh = 0;
	bResult = WriteFile( pPort->m_hComm,
						 pPort->m_szWriteBuffer,
						 pPort->m_nWriteSize,
						 &BytesSentA,
						 &pPort->m_ov );

	if ( !bResult )
	{
		DWORD dwError = GetLastError();

		switch ( dwError )
		{
			case ERROR_IO_PENDING:
				{
					bWrite = FALSE;
					break;
				}

			default:
				{
					pPort->ProcessErrorMessage( "WriteFile()" );
					break;
				}
		}
	}

	if ( !bWrite )
	{
		bResult = GetOverlappedResult( pPort->m_hComm,   // Handle to COMM port
									   &pPort->m_ov,     // Overlapped structure
									   &BytesSentB,      // Stores number of bytes sent
									   TRUE );           // Wait flag

		// deal with the error code
		if ( !bResult )
		{
			pPort->ProcessErrorMessage( "GetOverlappedResult() in WriteFile()" );
		}
	}

	if ( pPort->m_nWriteSize == sizeof( char ) )
	{
		FlushFileBuffers( pPort->m_hComm );
	}

	assert ( (BytesSentA + BytesSentB) == pPort->m_nWriteSize );
	pPort->m_nWriteSize = 0;
	LeaveCriticalSection( &pPort->m_csCommunicationSync );
	return (BytesSentA + BytesSentB);
}

void CSerialPort::ReceiveChar( CSerialPort *pPort )
{
	BOOL  bResult = TRUE;
	DWORD dwError = 0;
	DWORD BytesRead = 0;
	COMSTAT comstat;
	unsigned char RXBuff;

	while ( pPort->m_bThreadAlive )
	{
		if ( WaitForSingleObject( pPort->m_hShutdownEvent, 0 ) == WAIT_OBJECT_0 )
		{
			break;
		}

		// Gain ownership of the comm port critical section.
		// This process guarantees no other part of this program
		// is using the port object.
		EnterCriticalSection( &pPort->m_csCommunicationSync );
		// ClearCommError() will update the COMSTAT structure and
		// clear any other errors.
		bResult = ClearCommError( pPort->m_hComm, &dwError, &comstat );
		LeaveCriticalSection( &pPort->m_csCommunicationSync );

		// start forever loop.  I use this type of loop because I
		// do not know at runtime how many loops this will have to
		// run. My solution is to start a forever loop and to
		// break out of it when I have processed all of the
		// data available.  Be careful with this approach and
		// be sure your loop will exit.
		// My reasons for this are not as clear in this sample
		// as it is in my production code, but I have found this
		// solutiion to be the most efficient way to do this.

		if ( comstat.cbInQue == 0 )
		{
			break; // break out when all bytes have been read
		}

		EnterCriticalSection( &pPort->m_csCommunicationSync );
		BOOL bRead = TRUE;
		bResult = ReadFile( pPort->m_hComm,      // Handle to COMM port
							&RXBuff,             // RX Buffer Pointer
							sizeof( RXBuff ),    // Read one byte
							&BytesRead,          // Stores number of bytes read
							&pPort->m_ov );      // pointer to the m_ov structure

		if ( !bResult )
		{
			switch ( dwError = GetLastError() )
			{
				case ERROR_IO_PENDING:
					{
						// asynchronous i/o is still in progress
						// Proceed on to GetOverlappedResults();
						bRead = FALSE;
						break;
					}

				default:
					{
						pPort->ProcessErrorMessage( "ReadFile()" );
						break;
					}
			}
		}

		if ( !bRead )
		{
			bResult = GetOverlappedResult( pPort->m_hComm,   // Handle to COMM port
										   &pPort->m_ov,     // Overlapped structure
										   &BytesRead,       // Stores number of bytes read
										   TRUE );           // Wait flag

			if ( !bResult )
			{
				pPort->ProcessErrorMessage( "GetOverlappedResult() in ReadFile()" );
			}
		}

		LeaveCriticalSection( &pPort->m_csCommunicationSync );

		if ( bResult && ( BytesRead > 0 ) )
		{
			::PostMessage( pPort->m_pOwner, SERIAL_PORT_MESSAGE, ( WPARAM ) RXBuff, ( LPARAM )  EV_RXCHAR );
		}
	}
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

	m_dcb = *dcb;

	if ( SetCommState( m_hComm, &m_dcb ) == 0 )
	{
		ProcessErrorMessage( "SetCommState()" );
		ret = FALSE;
	}

	return ret;
}

BOOL CSerialPort::IsOpen()
{
	return m_hComm != INVALID_HANDLE_VALUE;
}

void CSerialPort::Close()
{
	if ( m_hShutdownEvent != NULL && m_bThreadAlive )
	{
		do
		{
			SetEvent( m_hShutdownEvent );
			::Sleep( 0 );
		}
		while ( m_bThreadAlive );
	}

	if ( m_hComm != INVALID_HANDLE_VALUE )
	{
		CloseHandle( m_hComm );
		m_hComm = INVALID_HANDLE_VALUE;
	}

	if ( m_hShutdownEvent != NULL )
	{
		ResetEvent( m_hShutdownEvent );
		m_hShutdownEvent = NULL;
	}

	if ( m_ov.hEvent != NULL )
	{
		ResetEvent( m_ov.hEvent );
		m_ov.hEvent = NULL;
	}

	if ( m_hWriteEvent != NULL )
	{
		ResetEvent( m_hWriteEvent );
		m_hWriteEvent = NULL;
	}

	memset( m_hEventArray, 0, sizeof( m_hEventArray ) );

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
}

void CSerialPort::Write( char *Buffer )
{
	int nSize = strlen( Buffer );
	assert( m_hComm != INVALID_HANDLE_VALUE );
	assert( Buffer != NULL );
	assert( nSize > 0 );
	EnterCriticalSection( &m_csCommunicationSync );
	assert( ( m_nWriteSize + nSize ) < ( int )m_nWriteBufferSize );
	strcpy( m_szWriteBuffer + m_nWriteSize, Buffer );
	m_nWriteSize += nSize;
	SetEvent( m_hWriteEvent );
	LeaveCriticalSection( &m_csCommunicationSync );
}

void CSerialPort::Write( void *Buffer, int nSize )
{
	assert( m_hComm != INVALID_HANDLE_VALUE );
	assert( Buffer != NULL );
	assert( nSize > 0 );
	EnterCriticalSection( &m_csCommunicationSync );
	assert( ( m_nWriteSize + nSize ) < ( int )m_nWriteBufferSize );
	memcpy( m_szWriteBuffer + m_nWriteSize, Buffer, nSize );
	m_nWriteSize += nSize;
	SetEvent( m_hWriteEvent );
	LeaveCriticalSection( &m_csCommunicationSync );
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
			MessageBox( m_pOwner, _T( "Failed to query registry!" ), m_szTitle, MB_ICONERROR );
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
