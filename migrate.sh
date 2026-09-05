#!/bin/bash
# ===== MIGRATION SCRIPT: 45.149.78.222 -> 5.159.49.12 =====
# Run this ON THE NEW SERVER (5.159.49.12)
set -e
OLD="root@45.149.78.222"
APP="/opt/salon-accounting"

echo "=== 1) Installing packages ==="
export DEBIAN_FRONTEND=noninteractive
apt-get update -y
apt-get install -y python3 python3-pip python3-venv nginx git curl ufw rsync

echo "=== 2) Pulling code from old server (includes ALL latest fixes) ==="
if [ -d "$APP" ]; then rm -rf "$APP"; fi
git clone https://github.com/pelakee13/salon-accounting.git "$APP"

echo "=== 3) Copying DATA from old server (customers, invoices, users) ==="
rsync -az -e "ssh -o StrictHostKeyChecking=no" $OLD:$APP/data/ $APP/data/
ls -la $APP/data/ 2>/dev/null || echo "(data dir empty?)"

echo "=== 4) Python venv + dependencies ==="
cd "$APP"
python3 -m venv venv
./venv/bin/pip install -q --upgrade pip
./venv/bin/pip install -q -r requirements.txt gunicorn

echo "=== 5) Systemd service ==="
cat > /etc/systemd/system/salon.service <<'EOF'
[Unit]
Description=Salon Accounting Gunicorn
After=network.target

[Service]
User=root
WorkingDirectory=/opt/salon-accounting
ExecStart=/opt/salon-accounting/venv/bin/gunicorn -w 3 -b 127.0.0.1:8000 web_app:app
Restart=always

[Install]
WantedBy=multi-user.target
EOF
systemctl daemon-reload
systemctl enable --now salon
sleep 3
systemctl is-active salon

echo "=== 6) Nginx ==="
cat > /etc/nginx/sites-available/salon <<'EOF'
server {
    listen 80;
    server_name _;
    client_max_body_size 20M;
    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
EOF
ln -sf /etc/nginx/sites-available/salon /etc/nginx/sites-enabled/salon
rm -f /etc/nginx/sites-enabled/default
nginx -t && systemctl restart nginx

echo "=== 7) Firewall ==="
ufw allow 22/tcp
ufw allow 80/tcp
yes | ufw enable

echo "=== 8) Local test ==="
sleep 2
curl -s -o /dev/null -w "root: %{http_code}\n" http://127.0.0.1/
curl -s -o /dev/null -w "login: %{http_code}\n" http://127.0.0.1/login

echo "=== DONE! Migration complete. ==="
echo "Test from browser: http://5.159.49.12"
