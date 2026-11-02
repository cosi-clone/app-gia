"""Vercel PDF function. Uses only Python standard libraries."""
import io
import ipaddress
import json
import socket
import time
import urllib.error
import urllib.parse
import urllib.request
from http.server import BaseHTTPRequestHandler

LIMIT = 4_000_000


def public_url(value):
    if len(value) > 8192:
        raise ValueError('Link quá dài.')
    p = urllib.parse.urlsplit(value)
    if p.scheme not in ('http', 'https') or not p.hostname or p.username or p.password:
        raise ValueError('Chỉ hỗ trợ link HTTP/HTTPS công khai, không chứa thông tin đăng nhập.')
    if p.port not in (None, 80, 443):
        raise ValueError('Chỉ hỗ trợ cổng HTTP/HTTPS thông thường.')
    for item in socket.getaddrinfo(p.hostname, p.port or (443 if p.scheme == 'https' else 80), type=socket.SOCK_STREAM):
        if not ipaddress.ip_address(item[4][0]).is_global:
            raise ValueError('Không hỗ trợ link tới localhost hoặc mạng nội bộ.')
    return value


class PublicRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        public_url(newurl)
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def download_pdf(url):
    public_url(url)
    request = urllib.request.Request(url, headers={'User-Agent': 'ExcelReader/1.0', 'Accept': 'application/pdf', 'Accept-Encoding': 'identity'})
    # No inherited proxy environment or browser cookies.
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}), PublicRedirect())
    started = time.monotonic()
    with opener.open(request, timeout=10) as response:
        length = response.headers.get('Content-Length', '')
        if length.isdigit() and int(length) > LIMIT:
            raise ValueError('PDF lớn hơn 4 MB. Hãy tải file về máy rồi upload PDF.')
        output = io.BytesIO()
        while True:
            if time.monotonic() - started > 40:
                raise TimeoutError('Nguồn PDF phản hồi quá lâu.')
            chunk = response.read(64 * 1024)
            if not chunk:
                break
            if output.tell() + len(chunk) > LIMIT:
                raise ValueError('PDF lớn hơn 4 MB. Hãy tải file về máy rồi upload PDF.')
            output.write(chunk)
        data = output.getvalue()
    if b'%PDF-' not in data[:1024]:
        raise ValueError('Link không trả về PDF. Có thể nguồn yêu cầu đăng nhập hoặc link đã hết hạn.')
    return data


class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        parsed = urllib.parse.urlsplit(self.path)
        if self.headers.get('Sec-Fetch-Site') == 'cross-site':
            return self.error_json(403, 'Chỉ cho phép yêu cầu từ giao diện ứng dụng.')
        try:
            url = urllib.parse.parse_qs(parsed.query).get('url', [''])[0]
            data = download_pdf(url)
            self.send_response(200)
            self.send_header('Content-Type', 'application/pdf')
            self.send_header('Content-Length', str(len(data)))
            self.send_header('Cache-Control', 'no-store')
            self.end_headers()
            self.wfile.write(data)
        except urllib.error.HTTPError as error:
            self.error_json(502, f'Nguồn PDF trả HTTP {error.code}.')
        except (TimeoutError, socket.timeout):
            self.error_json(504, 'Nguồn PDF phản hồi quá lâu. Thử lại sau hoặc tải file thủ công.')
        except (ValueError, urllib.error.URLError, OSError) as error:
            self.error_json(400, str(error))

    def error_json(self, status, message):
        data = json.dumps({'error': message}, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(data)))
        self.send_header('Cache-Control', 'no-store')
        self.end_headers()
        try:
            self.wfile.write(data)
        except (BrokenPipeError, ConnectionResetError):
            pass


