# MSB Windows Update Management Service

A simple Windows Service for cataloging, managing and installing 
windows updates.

## More Information
The service runs several timers that periodically execute processes on the local server. 
The main timer has the server check for available Windows Updates. Any available updates are automatically downloaded -- but not installed.
These updates are cataloged and can be viewed in the IT Dashboard.

![Approve Updates](https://github.com/Gimcrack/msb-windows-update-management/raw/master/images/ApproveUpdates.png "Approve Updates")

## Installation

Installation is performed using the [.NET InstallUtil.exe Installer Tool](https://msdn.microsoft.com/en-us/library/50614e95%28v=vs.110%29.aspx "Help Online")

## Usage

1. Install the Service.
2. Everything else is handled by the IT Dashboard.

## History

v1.0 Initial Commit

## Credits

Jeremy Bloomstrom Author

## License

MIT