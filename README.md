# RedTool - Python Version

## Overview

RedTool is a Python-based terminal and configuration tool for BOLT devices, converted from the original .NET LBHH Red Tool. This application provides serial communication capabilities, device configuration management, and TEK key handling for BOLT hardware.

## Features

### Core Functionality
- **Serial Communication**: Connect to BOLT devices via COM ports
- **BOLT Protocol Support**: Automatic BOLT protocol connection and communication
- **Multi-slot Configuration**: Configure up to 4 slots with XML-based configuration files
- **TEK Key Management**: Load and manage AME and WFC TEK files
- **Real-time Device Monitoring**: Monitor device status, brick number, unit ID, and more

### Enhanced Python Features
- **Modern GUI**: Built with tkinter for cross-platform compatibility
- **Progress Tracking**: Visual progress bar for configuration operations
- **Enhanced Logging**: Comprehensive logging system with file output
- **Configuration Management**: Import/export configuration settings
- **Validation Tools**: Built-in validation for TEK files and XML configurations
- **Error Handling**: Robust error handling with detailed error messages

## Requirements

### System Requirements
- Python 3.7 or higher
- Windows/Linux/macOS (tested on Windows)

### Python Dependencies
```bash
pip install pyserial>=3.5
```

Or install from requirements file:
```bash
pip install -r redtool_requirements.txt
```

## Installation

1. **Clone or download** the RedTool files to your local machine
2. **Install dependencies**:
   ```bash
   pip install pyserial
   ```
3. **Run the application**:
   ```bash
   python redtool.py
   ```

## Configuration

### Initial Setup
1. **Connect Hardware**: Connect your BOLT device to a COM port
2. **Select TEK Files**: Use File menu to choose AME and/or WFC TEK files
3. **Load Slot Configurations**: Load XML configuration files for each slot you want to configure

### TEK Files
TEK files should be XML format containing device IDs and encryption keys:
```xml
<TekKeys>
  <Device>
    <ID>12345</ID>
    <TEK_1>ABCDEF1234567890...</TEK_1>
  </Device>
</TekKeys>
```

### Slot Configuration Files
XML files defining waveform parameters:
```xml
<HWIWaveformConfig>
  <name>Slot Configuration Name</name>
  <Waveform>WaveformName</Waveform>
  <prompt>BOLT></prompt>
  <NeedsKey>Needs Key:</NeedsKey>
  <TEKLoad>loadtek</TEKLoad>
  <TEKFile>AME.tek</TEKFile>
  <WFType>LYNX-6</WFType>
  <Commands>
    <Command Type="command1" NumArguments="0"/>
  </Commands>
</HWIWaveformConfig>
```

## Usage

### Basic Operation
1. **Connect to Device**:
   - Select COM port from dropdown
   - Click "Connect" button
   - Application will attempt BOLT protocol connection

2. **Load Configuration**:
   - Select TEK files via File menu
   - Load XML configuration files for desired slots
   - Click "Load TEK Keys and Configure BOLT"

3. **Monitor Device**:
   - View real-time device information in status panel
   - Use "Refresh Info" to update device status
   - Send manual commands via command input

### Menu Options

#### File Menu
- **Choose AME/WFC TEK File**: Select encryption key files
- **Export Configuration**: Save current device and configuration state
- **Import Settings**: Load previously saved configuration settings
- **Backup Configuration**: Create backup of current settings

#### Tools Menu
- **Clear All Slot Configurations**: Remove all loaded XML configurations
- **Validate TEK Files**: Check all loaded files for errors
- **Device Diagnostics**: Run comprehensive device status check

#### Help Menu
- **About**: Display application information

## Configuration File Format

The application saves settings in `redtool_config.json`:
```json
{
    "ame_tek_path": "path/to/ame.tek",
    "wfc_tek_path": "path/to/wfc.tek",
    "serial_port": "COM2",
    "baudrate": 19200,
    "auto_connect": true,
    "auto_scroll": true,
    "char_delay": 0.01
}
```

## Troubleshooting

### Common Issues

1. **COM Port Not Found**
   - Ensure BOLT device is connected
   - Check Windows Device Manager for available ports
   - Try different COM ports

2. **BOLT Connection Failed**
   - Verify correct baud rate (typically 19200 initially, 115200 after connection)
   - Check cable connections
   - Ensure device is powered on

3. **TEK File Errors**
   - Verify XML format is correct
   - Check file permissions
   - Ensure device ID exists in TEK file

4. **Configuration Timeout**
   - Increase timeout values in XML
   - Check device responsiveness
   - Verify XML prompt strings match device output

### Debug Mode
Enable detailed logging by modifying the logging level in the code:
```python
logging.basicConfig(level=logging.DEBUG)
```

## Differences from .NET Version

### Improvements
- **Cross-platform compatibility** (works on Windows, Linux, macOS)
- **Enhanced error handling** with detailed logging
- **Progress tracking** for long operations
- **Modern Python practices** with type hints and dataclasses
- **Configuration import/export** functionality
- **Built-in validation tools**

### Known Limitations
- **Excel COM integration** not available (TEK files must be XML format)
- **Some advanced .NET features** may not be directly ported
- **Windows-specific optimizations** may differ

## Development

### Project Structure
```
RedTool/
├── redtool.py              # Main application file
├── redtool_config.json     # Configuration file
├── redtool_requirements.txt # Python dependencies
├── test_redtool.py         # Test script
└── README.md               # This file
```

### Testing
Run the test script to verify basic functionality:
```bash
python test_redtool.py
```

### Logging
Application logs are written to console and can be redirected to file:
```bash
python redtool.py > redtool.log 2>&1
```

## Support

For issues or questions:
1. Check the troubleshooting section above
2. Review log files for error details
3. Verify hardware connections and configuration files

## License

This software is converted from the original .NET LBHH Red Tool. Please refer to original licensing terms.

## Version History

### Version 2.0.0 (Python Conversion)
- Complete conversion from .NET to Python
- Enhanced GUI with progress tracking
- Improved error handling and logging
- Cross-platform compatibility
- Configuration import/export features
- Built-in validation tools

### Original .NET Version
- Windows Forms-based GUI
- Serial communication with BOLT devices
- XML-based configuration management
- TEK key handling
