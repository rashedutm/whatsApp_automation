# WhatsApp Bulk Message Automation

A Python-based automation tool for sending bulk messages with images to WhatsApp contacts using Selenium WebDriver.

## Features

- **Bulk Message Sending**: Send messages to multiple contacts from an Excel file
- **Image Support**: Send images with custom captions
- **Template Messages**: Use customizable message templates with contact name substitution
- **Progress Tracking**: Real-time feedback and logging of message delivery status
- **Emergency Stop**: Press `Ctrl+E` to safely stop automation at any time
- **Error Handling**: Robust error handling with retry mechanisms
- **Detailed Logging**: Comprehensive logging of all operations

## Prerequisites

- Python 3.7+
- Google Chrome browser
- ChromeDriver (compatible with your Chrome version)
- WhatsApp Web access

## Installation

1. Clone this repository:
```bash
git clone https://github.com/rashedutm/whatsApp_automation.git
cd whatsApp_automation
```

2. Install required Python packages:
```bash
pip install pandas selenium pyautogui pyperclip openpyxl keyboard
```

3. Download ChromeDriver:
   - Visit [ChromeDriver Downloads](https://chromedriver.chromium.org/)
   - Download the version compatible with your Chrome browser
   - Add ChromeDriver to your system PATH or place it in the project directory

## Setup

1. **Prepare your Excel file (`data.xlsx`)**:
   - Create an Excel file with columns for contact information
   - Include columns for names and phone numbers
   - Example format:
     ```
     Name                | Phone
     John Doe           | 1234567890
     Jane Smith         | 0987654321
     ```

2. **Configure your message template (`WHATSDRAFT.txt`)**:
   - Edit the `WHATSDRAFT.txt` file with your message template
   - Use `{name}` placeholder for contact name substitution
   - Example:
     ```
     Hi {name}!
     
     This is a test message from the automation tool.
     ```

3. **Add your image**:
   - Place your image file in the project directory
   - Update the image path in the code if needed

## Usage

1. **Start the automation**:
```bash
python whatsapp.py
```

2. **Follow the prompts**:
   - The script will guide you through the setup process
   - You'll need to scan the QR code for WhatsApp Web
   - The automation will start processing contacts from your Excel file

3. **Monitor progress**:
   - Check `feedback.txt` for detailed logs
   - The console will show real-time progress updates

4. **Emergency stop**:
   - Press `Ctrl+E` at any time to safely stop the automation
   - The script will complete the current contact and then stop

## File Structure

```
whatsApp_automation/
├── whatsapp.py          # Main automation script
├── data.xlsx            # Contact data (Excel file)
├── WHATSDRAFT.txt       # Message template
├── image.png            # Image to send with messages
├── feedback.txt         # Log file with operation results
└── README.md           # This file
```

## Configuration

### Excel File Format
The Excel file should contain:
- **Name column**: Contact names
- **Phone column**: Phone numbers (with country code, no spaces or special characters)

### Message Template
Edit `WHATSDRAFT.txt` to customize your message:
- Use `{name}` to include the contact's name
- Supports multi-line messages
- Keep messages concise for better delivery rates

## Safety Features

- **Rate Limiting**: Built-in delays between messages to avoid detection
- **Error Recovery**: Automatic retry mechanism for failed messages
- **Graceful Shutdown**: Safe stopping mechanism with `Ctrl+E`
- **Progress Preservation**: Logs all operations for tracking

## Troubleshooting

### Common Issues

1. **ChromeDriver not found**:
   - Ensure ChromeDriver is in your PATH or project directory
   - Check ChromeDriver version compatibility with your Chrome browser

2. **WhatsApp Web not loading**:
   - Check your internet connection
   - Ensure WhatsApp Web is accessible in your region
   - Try refreshing the page manually

3. **Messages not sending**:
   - Verify phone numbers are in correct format (with country code)
   - Check if contacts have WhatsApp accounts
   - Ensure you're not blocked by WhatsApp

4. **QR Code issues**:
   - Make sure your phone has internet connection
   - Try logging out and logging back into WhatsApp Web

### Error Logs
Check `feedback.txt` for detailed error information and operation logs.

## Disclaimer

This tool is for educational and personal use only. Please ensure you comply with:
- WhatsApp's Terms of Service
- Local laws and regulations regarding automated messaging
- Recipient consent for receiving messages

Use responsibly and avoid spam behavior to prevent account restrictions.

## Author

**Rashed--SE@UTM**

## License

This project is open source and available under the [MIT License](LICENSE).

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Support

For issues and questions, please open an issue on the GitHub repository.
