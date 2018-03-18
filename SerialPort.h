/*
**  FILENAME            SerialPort.h
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

#ifndef SERIAL_PORT_H
#define SERIAL_PORT_H

#define SERIAL_PORT_MAX             256UL                   /* http://digital.ni.com/public.nsf/allkb/F7A9002D7B8E31E7862568D6006BD10B */
#define MAX_VALUE_NAME              16383UL                 /* https://msdn.microsoft.com/en-us/library/ms724872(v=vs.85).aspx */
#define SERIAL_DEVICE_PREFIX        _T("COM")
#define WM_SERIAL_PORT_MESSAGE      _T("WM_SERIAL_PORT_MESSAGE_ID")

const static UINT  SERIAL_PORT_MESSAGE = ::RegisterWindowMessage( WM_SERIAL_PORT_MESSAGE );

class CSerialPort
{
    public:
        CSerialPort();
        virtual             ~CSerialPort();

        BOOL                Open( HWND  pPortOwner,
                                  UINT  port = 8,
                                  UINT  baud = 9600,
                                  BYTE  parity = NOPARITY,
                                  BYTE  databits = 8,
                                  BYTE  stopbits = ONESTOPBIT,
                                  DWORD dwCommEvents = EV_RXCHAR,
                                  UINT  nBufferSize = 4096,
                                  DWORD ReadIntervalTimeout = MAXDWORD,
                                  DWORD ReadTotalTimeoutMultiplier = 0,
                                  DWORD ReadTotalTimeoutConstant = 0,
                                  DWORD WriteTotalTimeoutMultiplier = 10,
                                  DWORD WriteTotalTimeoutConstant = 10 );
        void                Write( char *Buffer );
        void                Write( void *Buffer, int nSize );
        void                Close();

        DCB                 *GetDCB();
        BOOL                SetDCB( DCB *dcb );
        BOOL                IsOpen();
        void                EnumSerialPort( CComboBox &m_PortNO );

    protected:
        HANDLE              m_Thread;
        HANDLE              m_hComm;
        CRITICAL_SECTION    m_csCommunicationSync;
        COMMTIMEOUTS        m_CommTimeouts;
        DCB                 m_dcb;
        HWND                m_pOwner;
        volatile BOOL       m_bThreadAlive;
        volatile BOOL       m_bUserRequestClose;
        UINT                m_nPortNr;
        DWORD               m_dwCommEvents;
        DWORD               m_nWriteBufferSize;
        char                *m_szWriteBuffer;
        volatile int        m_nWriteSize;
        int                 m_nComArray[SERIAL_PORT_MAX + 1];

        static DWORD WINAPI CommThread( LPVOID pParam );
        static BOOL         ReceiveChar( CSerialPort *pPort );
        static UINT         WriteChar( CSerialPort *pPort );
        void                ProcessErrorMessage( char *ErrorText );
        BOOL                QueryRegistry( HKEY hKey );
};

#endif SERIAL_PORT_H
