IEEE Paper Formatter - website

1. Double-click run.bat  (installs Flask and python-docx, then starts the site)
2. Your browser opens http://localhost:5000
3. Upload a .docx, click "Format as IEEE", and the formatted file downloads.

Files:
  app.py             the web server (Flask)
  ieee_formatter.py  the formatting engine (IEEE rules are in the TEMPLATE block)
  templates/index.html  the web page
