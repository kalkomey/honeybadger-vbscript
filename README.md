# Honeybadger for VBScript (Classic ASP)

A client library for automatically sending VBScript errors to [Honeybadger](https://www.honeybadger.io/)

## Getting Started

1. Make sure your server has [`MSXML2.ServerXMLHTTP`](https://www.microsoft.com/en-us/download/details.aspx?id=4608) installed.
2. Copy `Honeybadger.asp` into your project's vendor or lib directory.
3. Create an `ErrorLogger.asp` file (see [`examples/ErrorLogger.asp`](https://github.com/kalkomey/honeybadger-vbscript/blob/master/examples/ErrorLogger.asp)).
4. Include `ErrorLogger.asp` in each ASP file on your website (or include it in a file already included by your ASP files).
5. Every exception is now logged in Honeybadger!

## License

Honeybadger-vbscript is released under the [MIT License](https://opensource.org/licenses/MIT).  See the file [`LICENSE`](https://github.com/kalkomey/honeybadger-vbscript/blob/master/LICENSE) for more information.
