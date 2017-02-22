CSerialPort
===========

####CSerialPort updated version

####Demo
####First step: 
     #include "SerialPort.h"

####Second step:
    CSerialPort port;
    port.Open( AfxGetApp()->m_pMainWnd->GetSafeHwnd(), 31 ); /* COM31 */
    port.GetDCB()->BaudRate = 9600;
    port.SetDCB(dcb);
    port.Write( _T("Hello World!") );
    port.Close();
    
####Event report register:
    ON_REGISTERED_MESSAGE( SERIAL_PORT_MESSAGE, &CMyDlg::OnPortMsg )

####Event process:
    LRESULT CMyDlg::OnPortMsg( WPARAM wParam, LPARAM lParam )
    {
	      LPARAM
		        EV_RXCHAR /* this event is receive, WPARAM is received char*/
		        EV_TXEMPTY
		        EV_CTS
		        EV_DSR
		        EV_RLSD
		        EV_RXFLAG
		        EV_BREAK
		        EV_ERR
		        EV_RING

	      WPARAM
		        char
    }

####10:19 2017/2/22

1. Fix warnings.
2. Support upto 256 physical serial port.

####12:56 2017/1/2

1. Add hot plug feature for usb to serial port cable.
2. Update code to satisfy Visual Studio 2015 static analysis.
3. Fix the event handles leakage issue in Close procedure.
