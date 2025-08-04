# Shopee Order Crawler

A powerful and intelligent web crawler for extracting order data from Shopee Seller Center. This tool automatically logs in, navigates through orders, and exports detailed order information to Excel format.

## 🚀 Features

- **Smart Login System**: Cookie-based authentication with automatic login persistence
- **Undetected Browser**: Uses undetected-chromedriver to avoid detection
- **Robust Element Interaction**: Multiple fallback strategies for clicking elements
- **Comprehensive Data Extraction**: Extracts order details, buyer info, delivery methods, and product specifications
- **Excel Export**: Automatically formats and exports data to Excel with proper styling
- **Error Handling**: Comprehensive error handling with retry mechanisms
- **Cross-Platform Support**: Works on Windows, macOS, and Linux

## 📋 Prerequisites

- Python 3.8 or higher
- Google Chrome browser
- Shopee Seller account

## 🛠️ Installation

1. **Clone the repository**

   ```bash
   git clone <your-repository-url>
   cd crawler
   ```

2. **Create a virtual environment (recommended)**

   ```bash
   python -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```

3. **Install dependencies**

   ```bash
   pip install -r requirements.txt
   ```

4. **Verify Chrome installation**
   - Ensure Google Chrome is installed on your system
   - The crawler will automatically download the appropriate ChromeDriver version

## 🎯 Usage

### Basic Usage

1. **Run the main crawler**

   ```bash
   python _shopee_order_crawler.py
   ```

2. **First-time setup**

   - The crawler will open a browser window
   - If not logged in, you'll be prompted to manually log in to Shopee Seller Center
   - After successful login, your session will be saved for future use

3. **Automatic data extraction**
   - The crawler will automatically navigate to the order page
   - Extract all available order data
   - Export results to Excel file in your Downloads folder

### Advanced Usage

#### Test the setup

```bash
python minimal_test.py
```

#### Run calculations utility (This is for shipment)

```bash
python Calculate.py
```

## 📊 Data Structure

The crawler extracts the following information for each order:

| Field           | Description                            |
| --------------- | -------------------------------------- |
| Date            | Order creation date                    |
| Buyer ID        | Customer username                      |
| Delivery Method | Shipping method (蝦皮, 7-11, OK, etc.) |
| Product Name    | Item name with emoji filtering         |
| Color           | Product color specification            |
| Size            | Product size specification             |
| Quantity        | Number of items ordered                |
| Total Price     | Order total amount                     |

## 🔧 Configuration

### Main Configuration

Edit the `config_dict` in `_shopee_order_crawler.py`:

```python
config_dict = {
    "advanced": {
        "verbose": True,      # Enable debug messages
        "headless": False     # Run browser in headless mode
    }
}
```

### Cookie Management

- Cookies are automatically saved to `shopee_cookies.pkl`
- To clear saved cookies: `python -c "from cookie_manager import CookieManager; CookieManager().clear_cookies()"`

## 📁 Project Structure

```
crawler/
├── _shopee_order_crawler.py    # Main crawler script
├── cookie_manager.py           # Cookie management system
├── minimal_test.py            # Simple test script
├── Calculate.py               # Utility calculation script
├── requirements.txt           # Python dependencies
├── .gitignore                # Git ignore rules
├── webdriver/                # ChromeDriver directory
│   └── chromedriver         # ChromeDriver binary
└── README.md                # This file
```

## 🛡️ Safety Features

### Element Clicking

The crawler implements multiple strategies for clicking elements:

1. Direct click
2. JavaScript click
3. ActionChains click
4. Overlay removal and click
5. Scroll to element before click

### Error Recovery

- Automatic retry mechanisms
- Graceful handling of missing elements
- Timeout management
- Browser crash recovery

## 🔍 Troubleshooting

### Common Issues

**1. ChromeDriver version mismatch**

```bash
# Clear ChromeDriver cache
rm -rf ~/.undetected_chromedriver
# Re-run the crawler
```

**2. Login issues**

```bash
# Clear saved cookies
python -c "from cookie_manager import CookieManager; CookieManager().clear_cookies()"
```

**3. Element not found errors**

- Ensure you're logged into Shopee Seller Center
- Check if the page structure has changed
- Try running in non-headless mode for debugging

**4. Permission errors**

```bash
# Make ChromeDriver executable (Linux/macOS)
chmod +x webdriver/chromedriver
```

### Debug Mode

Enable verbose logging by setting `"verbose": True` in the config dictionary.

## 📝 Output

The crawler generates an Excel file in your Downloads folder with the following naming convention:

- Format: `MM-DD.xlsx` (e.g., `12-25.xlsx`)
- Location: `~/Downloads/` directory
- Features: Merged cells, formatted styling, center alignment

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📄 License

This project is for educational and personal use. Please respect Shopee's terms of service and use responsibly.

## ⚠️ Disclaimer

- This tool is for educational purposes
- Use responsibly and in accordance with Shopee's terms of service
- The authors are not responsible for any misuse
- Respect rate limits and website policies

## 🆘 Support

If you encounter issues:

1. Check the troubleshooting section
2. Enable debug mode for detailed logs
3. Ensure all dependencies are properly installed
4. Verify your Shopee Seller account access

---

**Note**: This crawler is designed to work with the current Shopee Seller Center interface. If Shopee updates their website structure, the crawler may need updates to maintain compatibility.
