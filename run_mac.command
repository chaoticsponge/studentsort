#!/bin/sh
cd "$(dirname "$0")"

if [ -z "${SSL_CERT_FILE:-}" ] && [ -f /etc/ssl/cert.pem ]; then
  export SSL_CERT_FILE=/etc/ssl/cert.pem
fi

if [ -x /Library/Frameworks/Python.framework/Versions/3.11/bin/python3 ]; then
  /Library/Frameworks/Python.framework/Versions/3.11/bin/python3 exam_access_app.py
elif [ -x /Library/Frameworks/Python.framework/Versions/3.10/bin/python3 ]; then
  /Library/Frameworks/Python.framework/Versions/3.10/bin/python3 exam_access_app.py
else
  python3 exam_access_app.py
fi
