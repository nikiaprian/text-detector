#!/bin/bash

# Exit jika error dan gunakan variabel dengan aman
set -euo pipefail

# Inisialisasi variabel global
IONCUBE_SKIP=false

# Input domain
read -p "Masukkan nama domain (contoh: example.com): " DOMAIN_NAME

# Input password root MySQL
echo ""
read -p "Masukkan password root MySQL (akan digunakan untuk semua database): " MYSQL_ROOT_PASSWORD
echo ""

# Validasi password root MySQL
if [[ -z "$MYSQL_ROOT_PASSWORD" ]]; then
    echo "❌ Password root MySQL tidak boleh kosong"
    exit 1
fi

echo "✅ Password root MySQL telah diset"
echo ""

# Opsi untuk mengambil sertifikat SSL
read -p "Apakah ingin mengambil sertifikat SSL? (Y/n): " SSL_OPTION
SSL_OPTION=${SSL_OPTION:-Y}  # Default ke Y jika kosong

if [[ "$SSL_OPTION" =~ ^[Yy]$ ]]; then
    SSL_ENABLED=true
    echo "✅ SSL akan diaktifkan"
    echo ""
    read -p "Masukkan email untuk SSL (contoh: email@example.com): " EMAIL
    
    # Validasi email
    if [[ -z "$EMAIL" ]]; then
        echo "❌ Email tidak boleh kosong untuk SSL certificate"
        exit 1
    fi
else
    SSL_ENABLED=false
    echo "ℹ️  SSL dilewati, domain akan menggunakan HTTP"
    EMAIL=""  # Set email kosong jika SSL tidak diaktifkan
fi
echo ""

# Opsi untuk menginstall IonCube Loader
read -p "Apakah ingin menginstall IonCube Loader? (Y/n): " IONCUBE_OPTION
IONCUBE_OPTION=${IONCUBE_OPTION:-Y}  # Default ke Y jika kosong

if [[ "$IONCUBE_OPTION" =~ ^[Yy]$ ]]; then
    IONCUBE_ENABLED=true
    echo "✅ IonCube Loader akan diinstall"
else
    IONCUBE_ENABLED=false
    echo "ℹ️  IonCube Loader dilewati"
fi
echo ""

# Generate random password MySQL (20 karakter)
MYSQL_PASSWORD=$(openssl rand -base64 32 | tr -d "=+/" | cut -c1-20)

# Cek apakah domain sudah ada
if [ -d "/var/www/html/$DOMAIN_NAME" ]; then
    echo ""
    echo "⚠️  PERINGATAN: Domain '$DOMAIN_NAME' sudah ada!"
    echo "📋 Yang akan terjadi:"
    echo "   • WordPress files akan dihapus dan diganti"
    echo "   • Database akan dihapus dan dibuat ulang"
    echo "   • SSL certificate akan dipertahankan (jika ada)"
    echo "   • Konfigurasi Nginx akan diupdate"
    echo ""
    read -p "Apakah Anda yakin ingin melanjutkan? (ketik 'yes' untuk konfirmasi): " CONFIRM
    
    if [ "$CONFIRM" != "yes" ]; then
        echo "❌ Proses dibatalkan."
        exit 0
    fi
    echo ""
fi

DB_NAME="db_${DOMAIN_NAME//./}"
# Generate DB_USER with maximum 32 characters (MySQL limit)
# Use first 8 chars + hash of domain to ensure uniqueness while staying under limit
DOMAIN_CLEAN="${DOMAIN_NAME//./}"
DOMAIN_HASH=$(echo -n "$DOMAIN_NAME" | md5sum | cut -c1-8)
DB_USER="u_$(echo "$DOMAIN_CLEAN" | cut -c1-20)_$DOMAIN_HASH"
# Ensure DB_USER doesn't exceed 32 characters
if [ ${#DB_USER} -gt 32 ]; then
    DB_USER="u_$(echo "$DOMAIN_CLEAN" | cut -c1-15)_$DOMAIN_HASH"
fi
REDIS_PREFIX=$(echo "$DOMAIN_NAME" | tr -d '.' | cut -c1-8)
PHPMYADMIN_USER="phpmyadmin"
WP_CONFIG="/var/www/html/$DOMAIN_NAME/wp-config-sample.php"

# Update dan upgrade sistem
echo "🔄 Updating system..."
sudo apt update && sudo apt upgrade -y

# Install Nginx jika belum ada
if ! command -v nginx &> /dev/null; then
  echo "⚙️ Installing Nginx..."
  sudo apt install -y nginx
  echo "✅ Nginx installed successfully"
else
  echo "✅ Nginx sudah terpasang."
fi

# Install PHP-FPM dan ekstensi yang diperlukan
echo "🐘 Installing PHP-FPM dan ekstensi..."
sudo apt install -y php-fpm php-mysql php-mbstring php-zip php-gd php-json php-curl php-xml php-xmlrpc php-soap php-intl php-cli

# Deteksi versi PHP yang terinstal
echo "🔍 Mendeteksi versi PHP yang terinstal..."
PHP_VERSION=$(php -r 'echo PHP_MAJOR_VERSION.".".PHP_MINOR_VERSION;' 2>/dev/null || echo "8.1")
echo "✅ PHP versi $PHP_VERSION terdeteksi"

# Set path socket PHP-FPM berdasarkan versi
PHP_FPM_SOCK="/run/php/php${PHP_VERSION}-fpm.sock"
echo "📝 PHP-FPM socket: $PHP_FPM_SOCK"


# Siapkan direktori WordPress
echo "📁 Mempersiapkan direktori /var/www/html..."
sudo mkdir -p /var/www/html
cd /var/www/html

# Fungsi untuk mendapatkan versi dari file yang ada
get_existing_wordpress_version() {
    local file_path="$1"
    if [ -f "$file_path" ]; then
        # Extract version from filename: wordpress-6.8.3.tar.gz -> 6.8.3
        local filename=$(basename "$file_path")
        local version=$(echo "$filename" | sed -n 's/wordpress-\([0-9]\+\.[0-9]\+\.[0-9]\+\).*/\1/p')
        echo "$version"
    else
        echo ""
    fi
}

# Fungsi untuk membersihkan file WordPress lama (opsional)
cleanup_old_wordpress_files() {
    echo "🧹 Membersihkan file WordPress lama..."
    local current_file="$1"
    
    # Hapus file latest.tar.gz jika ada
    if [ -f "/var/www/html/latest.tar.gz" ]; then
        echo "🗑️  Menghapus file latest.tar.gz lama..."
        sudo rm -f "/var/www/html/latest.tar.gz"
    fi
    
    # Hapus file wordpress-*.tar.gz kecuali yang sedang digunakan
    for file in /var/www/html/wordpress-*.tar.gz; do
        if [ -f "$file" ] && [ "$file" != "$current_file" ]; then
            echo "🗑️  Menghapus file WordPress lama: $(basename "$file")"
            sudo rm -f "$file"
        fi
    done
}

# Download WordPress dengan pengecekan versi yang akurat
echo "🔍 Mengecek versi WordPress terbaru..."
LATEST_VERSION=$(curl -s https://api.wordpress.org/core/version-check/1.7/ | grep -o '"version":"[^"]*' | head -1 | cut -d'"' -f4)
WP_FILE="wordpress-${LATEST_VERSION}.tar.gz"
WP_FILE_PATH="/var/www/html/${WP_FILE}"

echo "📋 Versi WordPress terbaru: $LATEST_VERSION"

# Cek apakah file dengan versi terbaru sudah ada
EXISTING_VERSION=$(get_existing_wordpress_version "$WP_FILE_PATH")

if [ "$EXISTING_VERSION" = "$LATEST_VERSION" ] && [ -f "$WP_FILE_PATH" ]; then
  echo "✅ WordPress $LATEST_VERSION sudah tersedia dan versinya sesuai"
  echo "📁 Menggunakan file: $WP_FILE"
  
  # Bersihkan file WordPress lama
  cleanup_old_wordpress_files "$WP_FILE_PATH"
else
  if [ -n "$EXISTING_VERSION" ]; then
    echo "⚠️  File WordPress ada tapi versi berbeda:"
    echo "   📦 File existing: $EXISTING_VERSION"
    echo "   🆕 Versi terbaru: $LATEST_VERSION"
    echo "   🔄 Akan download versi terbaru..."
  else
    echo "📦 File WordPress $LATEST_VERSION belum ada, akan download..."
  fi
  
  echo "⬇️ Mengunduh WordPress $LATEST_VERSION..."
  sudo wget "https://wordpress.org/${WP_FILE}" -O "$WP_FILE_PATH"
  
  # Verifikasi download berhasil
  if [ -f "$WP_FILE_PATH" ]; then
    echo "✅ WordPress $LATEST_VERSION berhasil diunduh"
    
    # Bersihkan file WordPress lama setelah download berhasil
    cleanup_old_wordpress_files "$WP_FILE_PATH"
  else
    echo "❌ Gagal mengunduh WordPress $LATEST_VERSION"
    exit 1
  fi
fi

# Ekstrak dan pindahkan
echo "📦 Mengekstrak WordPress $LATEST_VERSION..."
sudo tar -xzf "$WP_FILE_PATH"

# Hapus folder domain lama jika ada
if [ -d "/var/www/html/$DOMAIN_NAME" ]; then
  echo "⚠️ Folder /var/www/html/$DOMAIN_NAME sudah ada. Menghapus..."
  sudo rm -rf "/var/www/html/$DOMAIN_NAME"
fi

echo "🚚 Memindahkan WordPress ke /var/www/html/$DOMAIN_NAME..."
sudo mv /var/www/html/wordpress "/var/www/html/$DOMAIN_NAME"

# Konfigurasi WordPress
echo "⚙️ Mengonfigurasi WordPress..."

# Update database credentials di wp-config-sample.php
echo "📝 Mengupdate database credentials..."
sudo sed -i "s/define( 'DB_NAME', '.*' );/define( 'DB_NAME', '$DB_NAME' );/" "$WP_CONFIG"
sudo sed -i "s/define( 'DB_USER', '.*' );/define( 'DB_USER', '$DB_USER' );/" "$WP_CONFIG"
sudo sed -i "s/define( 'DB_PASSWORD', '.*' );/define( 'DB_PASSWORD', '$MYSQL_PASSWORD' );/" "$WP_CONFIG"
sudo sed -i "s/define( 'DB_HOST', '.*' );/define( 'DB_HOST', 'localhost' );/" "$WP_CONFIG"
echo "✅ Database credentials diupdate:"
echo "   • DB_NAME: $DB_NAME"
echo "   • DB_USER: $DB_USER"
echo "   • DB_PASSWORD: $MYSQL_PASSWORD"
echo "   • DB_HOST: localhost"

# Tambahkan baris define di bawah DB_COLLATE
if ! grep -q "FS_METHOD" "$WP_CONFIG"; then
  sudo sed -i "/define( 'DB_COLLATE', '' );/a define('FS_METHOD', 'direct');\ndefine('WP_REDIS_UNIQUE_PREFIX', '${REDIS_PREFIX}');\n" "$WP_CONFIG"
  echo "✅ Ditambahkan define('FS_METHOD', 'direct') dan WP_REDIS_UNIQUE_PREFIX='$REDIS_PREFIX' ke $WP_CONFIG"
else
  echo "ℹ️ Konfigurasi Redis sudah ada di $WP_CONFIG, dilewati."
fi

# Rename wp-config-sample.php menjadi wp-config.php
echo "📝 Mengaktifkan wp-config.php..."
if [ -f "/var/www/html/$DOMAIN_NAME/wp-config-sample.php" ]; then
  sudo mv /var/www/html/$DOMAIN_NAME/wp-config-sample.php /var/www/html/$DOMAIN_NAME/wp-config.php
  echo "✅ wp-config-sample.php berhasil diubah menjadi wp-config.php"
elif [ -f "/var/www/html/$DOMAIN_NAME/wp-config.php" ]; then
  echo "✅ wp-config.php sudah ada"
else
  echo "⚠️  Warning: wp-config.php tidak ditemukan"
fi

# Install Certbot dan Nginx plugin jika belum ada
if ! command -v certbot &> /dev/null; then
  echo "🔧 Menginstal Certbot dan Nginx plugin..."
  sudo apt install certbot python3-certbot-nginx -y
else
  echo "✅ Certbot sudah terpasang."
  # Pastikan nginx plugin terinstal
  if ! dpkg -l | grep -q python3-certbot-nginx; then
    echo "📦 Menginstal Certbot Nginx plugin..."
    sudo apt install python3-certbot-nginx -y
  fi
fi

# Install dan setup MySQL
echo "🛠️ Menginstal MySQL Server..."
sudo apt install -y mysql-server


# Set akun admin MySQL khusus untuk phpMyAdmin
echo "🔐 Menyiapkan akun MySQL khusus phpMyAdmin..."
sudo mysql <<EOF
CREATE USER IF NOT EXISTS '$PHPMYADMIN_USER'@'127.0.0.1' IDENTIFIED BY '$MYSQL_ROOT_PASSWORD';
ALTER USER '$PHPMYADMIN_USER'@'127.0.0.1' IDENTIFIED BY '$MYSQL_ROOT_PASSWORD';
GRANT ALL PRIVILEGES ON *.* TO '$PHPMYADMIN_USER'@'127.0.0.1' WITH GRANT OPTION;
FLUSH PRIVILEGES;
EOF

# Pastikan root@localhost tetap tanpa password (hanya akses lokal via socket)
sudo mysql -e "ALTER USER 'root'@'localhost' IDENTIFIED WITH auth_socket;" 2>/dev/null || true

sudo mysql -e "FLUSH PRIVILEGES;" 2>/dev/null || true
echo "✅ Akun phpMyAdmin ($PHPMYADMIN_USER@127.0.0.1) siap"
echo "✅ Root@localhost tetap tanpa password (auth_socket)"

# Konfigurasi keamanan MySQL secara manual (non-interaktif)
echo "🛡️ Mengonfigurasi keamanan MySQL..."
sudo mysql <<EOF
DELETE FROM mysql.user WHERE User='';
DELETE FROM mysql.user WHERE User='root' AND Host NOT IN ('localhost', 'localhost', '::1', '%');
DROP DATABASE IF EXISTS test;
DELETE FROM mysql.db WHERE Db='test' OR Db='test\\_%';
FLUSH PRIVILEGES;
EOF
echo "✅ MySQL security configuration selesai"

echo "🧩 Setup database dan user..."
sudo mysql <<EOF
DROP DATABASE IF EXISTS $DB_NAME;
DROP USER IF EXISTS '$DB_USER'@'localhost';
CREATE DATABASE $DB_NAME DEFAULT CHARACTER SET utf8 COLLATE utf8_unicode_ci;
CREATE USER '$DB_USER'@'localhost' IDENTIFIED BY '$MYSQL_PASSWORD';
GRANT ALL ON $DB_NAME.* TO '$DB_USER'@'localhost';
FLUSH PRIVILEGES;
EOF

# Konfigurasi phpMyAdmin sebelum instalasi untuk menghindari prompt
echo "📦 Mengatur konfigurasi phpMyAdmin sebelum instalasi..."
export DEBIAN_FRONTEND=noninteractive
sudo bash -c "echo 'phpmyadmin phpmyadmin/reconfigure-webserver multiselect' | debconf-set-selections"
sudo bash -c "echo 'phpmyadmin phpmyadmin/dbconfig-install boolean true' | debconf-set-selections"
sudo bash -c "echo 'phpmyadmin phpmyadmin/mysql/admin-pass password ' | debconf-set-selections"
sudo bash -c "echo 'phpmyadmin phpmyadmin/mysql/app-pass password ' | debconf-set-selections"

echo "📦 Menginstal phpMyAdmin..."
sudo apt install -y phpmyadmin --no-install-recommends

# Konfigurasi phpMyAdmin agar memakai akun khusus
echo "🔧 Memastikan akun phpMyAdmin memiliki akses penuh..."
sudo mysql <<EOF
CREATE USER IF NOT EXISTS '$PHPMYADMIN_USER'@'127.0.0.1' IDENTIFIED BY '$MYSQL_ROOT_PASSWORD';
ALTER USER '$PHPMYADMIN_USER'@'127.0.0.1' IDENTIFIED BY '$MYSQL_ROOT_PASSWORD';
GRANT ALL PRIVILEGES ON *.* TO '$PHPMYADMIN_USER'@'127.0.0.1' WITH GRANT OPTION;

CREATE USER IF NOT EXISTS 'root'@'localhost' IDENTIFIED WITH auth_socket;
ALTER USER 'root'@'localhost' IDENTIFIED WITH auth_socket;

FLUSH PRIVILEGES;
EOF

# Buat database phpmyadmin jika belum ada
sudo mysql <<EOF
CREATE DATABASE IF NOT EXISTS phpmyadmin;
GRANT ALL PRIVILEGES ON phpmyadmin.* TO '$PHPMYADMIN_USER'@'127.0.0.1';
FLUSH PRIVILEGES;
EOF

# Update konfigurasi phpMyAdmin
echo "📝 Mengupdate konfigurasi phpMyAdmin..."
sudo tee /etc/phpmyadmin/config.inc.php > /dev/null <<EOF
<?php
\$cfg['blowfish_secret'] = '$(openssl rand -base64 32)';

// Server configuration
\$i = 0;
\$i++;
\$cfg['Servers'][\$i]['auth_type'] = 'cookie';
\$cfg['Servers'][\$i]['host'] = '127.0.0.1';
\$cfg['Servers'][\$i]['compress'] = false;
\$cfg['Servers'][\$i]['AllowNoPassword'] = false;
\$cfg['Servers'][\$i]['AllowRoot'] = false;
\$cfg['Servers'][\$i]['user'] = '$PHPMYADMIN_USER';
\$cfg['Servers'][\$i]['password'] = '$MYSQL_ROOT_PASSWORD';
\$cfg['Servers'][\$i]['connect_type'] = 'tcp';
\$cfg['Servers'][\$i]['extension'] = 'mysqli';
\$cfg['Servers'][\$i]['port'] = 3306;

// General configuration
\$cfg['UploadDir'] = '';
\$cfg['SaveDir'] = '';
\$cfg['TempDir'] = '/tmp';
\$cfg['CheckConfigurationPermissions'] = false;
\$cfg['DefaultLang'] = 'en';
\$cfg['ServerDefault'] = 1;

// Security
\$cfg['ForceSSL'] = false;
\$cfg['AllowArbitraryServer'] = false;
\$cfg['LoginCookieValidity'] = 1440;

// UI
\$cfg['ThemeDefault'] = 'pmahomme';
\$cfg['DefaultTabServer'] = 'welcome';
\$cfg['DefaultTabDatabase'] = 'structure';
\$cfg['DefaultTabTable'] = 'browse';
EOF

# Set permission yang benar
sudo chmod 644 /etc/phpmyadmin/config.inc.php
sudo chown root:root /etc/phpmyadmin/config.inc.php

# Pastikan MySQL bisa diakses dari PHP
echo "🔧 Mengkonfigurasi MySQL untuk akses dari PHP..."
sudo mysql <<EOF
-- Pastikan bind-address memungkinkan koneksi lokal
-- (Ini akan di-handle oleh konfigurasi MySQL default)
FLUSH PRIVILEGES;
EOF

# Test koneksi MySQL dari PHP
echo "🧪 Testing koneksi MySQL dari PHP..."
php -r "
// Test dengan root@localhost (tanpa password)
\$link1 = mysqli_connect('localhost', 'root', '');
if (\$link1) {
    echo '✅ Koneksi MySQL root@localhost (CLI) berhasil\n';
    mysqli_close(\$link1);
} else {
    echo '❌ Koneksi MySQL root@localhost gagal: ' . mysqli_connect_error() . '\n';
}

// Test dengan akun phpMyAdmin (dengan password)
\$link2 = mysqli_connect('127.0.0.1', '$PHPMYADMIN_USER', '$MYSQL_ROOT_PASSWORD');
if (\$link2) {
    echo '✅ Koneksi MySQL user phpMyAdmin berhasil\n';
    mysqli_close(\$link2);
} else {
    echo '❌ Koneksi MySQL user phpMyAdmin gagal: ' . mysqli_connect_error() . '\n';
}
"

# Install Redis Server
echo "🔴 Menginstal Redis Server..."
sudo apt install -y redis-server

# Konfigurasi Firewall
echo "🛡️ Mengatur firewall..."
sudo ufw allow 'Nginx Full'
sudo ufw --force enable
sudo ufw allow 22/tcp
sudo ufw allow 51623/tcp

# Set hak akses
echo "🔐 Mengatur hak akses direktori dan file..."
# Set ownership
sudo chown -R www-data:www-data /var/www/html/$DOMAIN_NAME

# Set folder permissions to 755
echo "   📁 Setting folder permissions to 755..."
sudo find /var/www/html/$DOMAIN_NAME -type d -exec chmod 755 {} \;

# Set file permissions to 644
echo "   📄 Setting file permissions to 644..."
sudo find /var/www/html/$DOMAIN_NAME -type f -exec chmod 644 {} \;

# Set wp-config.php permission to 440 (read-only for owner and group)
if [ -f "/var/www/html/$DOMAIN_NAME/wp-config.php" ]; then
    sudo chmod 440 /var/www/html/$DOMAIN_NAME/wp-config.php
    sudo chown www-data:www-data /var/www/html/$DOMAIN_NAME/wp-config.php
    echo "   🔒 wp-config.php: 440 (read-only)"
else
    echo "   ⚠️  wp-config.php tidak ditemukan"
fi

echo ""
echo "✅ Permissions Summary:"
echo "   • Owner: www-data:www-data"
echo "   • Folders: 755 (rwxr-xr-x)"
echo "   • Files: 644 (rw-r--r--)"
echo "   • wp-config.php: 440 (r--r-----)"

# Setup Nginx configuration
echo "📝 Membuat konfigurasi Nginx untuk $DOMAIN_NAME..."

# Buat direktori sites-available dan sites-enabled jika belum ada
sudo mkdir -p /etc/nginx/sites-available
sudo mkdir -p /etc/nginx/sites-enabled

# Buat konfigurasi HTTP dulu (SSL akan ditambahkan oleh certbot --nginx)
echo "📝 Membuat konfigurasi HTTP untuk $DOMAIN_NAME..."
sudo tee /etc/nginx/sites-available/$DOMAIN_NAME > /dev/null <<EOF
server {
    listen 80;
    listen [::]:80;
    server_name $DOMAIN_NAME;

    root /var/www/html/$DOMAIN_NAME;
    index index.php index.html index.htm;

    client_max_body_size 800M;

    location / {
        try_files \$uri \$uri/ /index.php?\$args;
    }

    location ~ \.php$ {
        include snippets/fastcgi-php.conf;
        fastcgi_pass unix:$PHP_FPM_SOCK;
        fastcgi_param SCRIPT_FILENAME \$document_root\$fastcgi_script_name;
        include fastcgi_params;
    }

    # Security: Block suspicious files
    location = /images/toggige-arrow.jpg {
        return 404;
    }

    location ~ /toggige-arrow\.jpg$ {
        return 404;
    }

    location ~* toggige.*\.jpg$ {
        return 404;
    }

    # Security: Disable PHP execution in upload directories
    location ~* ^/(wp-content/uploads|images|media|files|assets/uploads)/ {
        location ~ \.php$ {
            deny all;
        }
        try_files \$uri =404;
    }

    # Security: Block access to sensitive WordPress files
    location ~* /(wp-config\.php|readme\.html|license\.txt|xmlrpc\.php) {
        deny all;
    }

    # Security: Block access to hidden files (EXCEPT .well-known for SSL)
    location ~ /\.(?!well-known).* {
        deny all;
        access_log off;
        log_not_found off;
    }

    location ~* \.(ico|css|js|gif|jpg|jpeg|png|svg|woff|woff2|ttf|eot)$ {
        expires max;
        log_not_found off;
        access_log off;
    }

    location ~ /\.ht {
        deny all;
    }

    location = /favicon.ico {
        log_not_found off;
        access_log off;
    }

    location = /robots.txt {
        allow all;
        log_not_found off;
        access_log off;
    }

    gzip on;
    gzip_vary on;
    gzip_min_length 1000;
    gzip_comp_level 6;
    gzip_types text/plain text/css text/xml text/javascript application/x-javascript application/xml+rss application/javascript application/json;
}
EOF

# Enable site dengan symbolic link
echo "🔗 Mengaktifkan site configuration..."
sudo ln -sf /etc/nginx/sites-available/$DOMAIN_NAME /etc/nginx/sites-enabled/

# Buat konfigurasi phpMyAdmin di Nginx (port 51623)
echo "📝 Membuat konfigurasi Nginx untuk phpMyAdmin..."
sudo tee /etc/nginx/sites-available/phpmyadmin > /dev/null <<EOF
server {
    listen 51623;
    listen [::]:51623;
    
    root /usr/share/phpmyadmin;
    index index.php index.html index.htm;

    location / {
        try_files \$uri \$uri/ =404;
    }

    location ~ \.php$ {
        include snippets/fastcgi-php.conf;
        fastcgi_pass unix:$PHP_FPM_SOCK;
        fastcgi_param SCRIPT_FILENAME \$document_root\$fastcgi_script_name;
        include fastcgi_params;
    }

    location ~* \.(js|css|png|jpg|jpeg|gif|ico)$ {
        expires max;
        log_not_found off;
    }

    gzip on;
    gzip_vary on;
    gzip_comp_level 6;
    gzip_types text/plain text/css application/json application/javascript text/xml application/xml application/xml+rss text/javascript;
}
EOF

# Enable phpMyAdmin site
sudo ln -sf /etc/nginx/sites-available/phpmyadmin /etc/nginx/sites-enabled/

# Test konfigurasi Nginx
echo "🧪 Testing Nginx configuration..."
if sudo nginx -t; then
    echo "✅ Nginx configuration valid"
else
    echo "❌ Nginx configuration error"
    exit 1
fi

# Enable dan start services
echo "🔄 Mengaktifkan dan menjalankan services..."
sudo systemctl daemon-reload
sudo systemctl enable nginx
sudo systemctl enable php${PHP_VERSION}-fpm
sudo systemctl restart php${PHP_VERSION}-fpm
sudo systemctl restart nginx

# Setup SSL dengan Certbot Nginx Plugin jika enabled
if [ "$SSL_ENABLED" = true ]; then
    echo ""
    echo "🔐 Mengaktifkan SSL dengan Certbot Nginx Plugin..."
    echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    
    if [ ! -d "/etc/letsencrypt/live/$DOMAIN_NAME" ]; then
        echo "🔐 Mendapatkan SSL certificate untuk $DOMAIN_NAME..."
        echo "ℹ️  Method: Certbot Nginx Plugin"
        
        # Gunakan certbot --nginx untuk auto-configure HTTPS
        if sudo certbot --nginx \
            --non-interactive \
            --agree-tos \
            --email "$EMAIL" \
            -d "$DOMAIN_NAME" \
            --redirect; then
            
            echo "✅ SSL certificate berhasil diperoleh dan dikonfigurasi!"
            echo "✅ HTTPS telah diaktifkan dengan auto-redirect dari HTTP"
            
            # Tambahkan SSL hardening tambahan
            echo "🔒 Menambahkan SSL optimization dan security headers..."
            
            # Backup config yang dibuat certbot
            sudo cp "/etc/nginx/sites-available/$DOMAIN_NAME" "/etc/nginx/sites-available/$DOMAIN_NAME.certbot-backup"
            
            # Tambahkan SSL optimization dan security headers
            if ! grep -q "ssl_session_cache" "/etc/nginx/sites-available/$DOMAIN_NAME"; then
                # Cari baris ssl_certificate dan tambahkan setelahnya
                sudo sed -i "/ssl_certificate_key/a \\\n    # SSL Optimization\n    ssl_session_cache shared:SSL:10m;\n    ssl_session_timeout 10m;\n    ssl_session_tickets off;\n    ssl_stapling on;\n    ssl_stapling_verify on;\n    resolver 8.8.8.8 8.8.4.4 valid=300s;\n    resolver_timeout 5s;\n\n    # Security Headers\n    add_header Strict-Transport-Security \"max-age=31536000; includeSubDomains\" always;\n    add_header X-Frame-Options \"SAMEORIGIN\" always;\n    add_header X-Content-Type-Options \"nosniff\" always;\n    add_header X-XSS-Protection \"1; mode=block\" always;" "/etc/nginx/sites-available/$DOMAIN_NAME"
            fi
            
            # Validasi security blocks masih ada
            echo "🔍 Validasi security blocks..."
            MISSING_BLOCKS=()
            
            if ! grep -q "toggige-arrow" "/etc/nginx/sites-available/$DOMAIN_NAME"; then
                MISSING_BLOCKS+=("toggige-arrow block")
            fi
            
            if ! grep -q "wp-content/uploads" "/etc/nginx/sites-available/$DOMAIN_NAME"; then
                MISSING_BLOCKS+=("PHP in uploads block")
            fi
            
            if ! grep -q "wp-config" "/etc/nginx/sites-available/$DOMAIN_NAME"; then
                MISSING_BLOCKS+=("sensitive files block")
            fi
            
            if [ ${#MISSING_BLOCKS[@]} -gt 0 ]; then
                echo "⚠️  WARNING: Certbot removed some security blocks:"
                for block in "${MISSING_BLOCKS[@]}"; do
                    echo "   ❌ $block"
                done
                echo ""
                echo "💡 Security blocks were in HTTP config but not transferred to HTTPS"
                echo "📁 Review config: /etc/nginx/sites-available/$DOMAIN_NAME"
                echo "📋 Original config: /etc/nginx/sites-available/$DOMAIN_NAME.certbot-backup"
            else
                echo "✅ All security blocks preserved"
            fi
            
            # Test dan reload
            if sudo nginx -t; then
                sudo systemctl reload nginx
                echo "✅ SSL optimization & security headers applied"
            else
                echo "❌ Config error after SSL optimization"
                echo "⚠️  Restoring certbot-backup..."
                sudo cp "/etc/nginx/sites-available/$DOMAIN_NAME.certbot-backup" "/etc/nginx/sites-available/$DOMAIN_NAME"
                sudo systemctl reload nginx
            fi
        else
            echo "❌ Gagal mendapatkan SSL certificate"
            echo "⚠️  Domain akan tetap menggunakan HTTP"
            echo "💡 Pastikan:"
            echo "   • Domain $DOMAIN_NAME sudah pointing ke server IP ini"
            echo "   • Port 80 & 443 terbuka di firewall"
            echo "   • Tidak ada masalah DNS"
        fi
    else
        echo "📄 SSL certificate untuk $DOMAIN_NAME sudah ada"
        echo "🔄 Mengupdate konfigurasi Nginx untuk menggunakan SSL..."
        
        # Gunakan certbot untuk update config dengan SSL existing
        if sudo certbot --nginx \
            --non-interactive \
            --cert-name "$DOMAIN_NAME" \
            --redirect; then
            echo "✅ Konfigurasi Nginx telah diupdate untuk menggunakan SSL"
        fi
    fi
else
    echo ""
    echo "ℹ️  SSL dilewati - domain menggunakan HTTP saja"
fi

# Cek status services
echo ""
echo "📊 Status Nginx service:"
sudo systemctl status nginx --no-pager -l

echo ""
echo "📊 Status PHP-FPM service:"
sudo systemctl status php${PHP_VERSION}-fpm --no-pager -l

# Info akhir
echo ""
echo "✅ Instalasi selesai!"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""
echo "📁 File Location:"
echo "   • WordPress Path: /var/www/html/$DOMAIN_NAME"
echo "   • Nginx Config: /etc/nginx/sites-available/$DOMAIN_NAME"
echo ""
echo "🗄️  Database Information:"
echo "   • DB Name: $DB_NAME"
echo "   • DB User: $DB_USER"
echo "   • DB Pass: $MYSQL_PASSWORD"
echo "   • DB Host: localhost"
echo "   • Redis Prefix: $REDIS_PREFIX"
echo "   • MySQL Root Password: $MYSQL_ROOT_PASSWORD"
echo "   • phpMyAdmin User: $PHPMYADMIN_USER@127.0.0.1"
echo ""
echo "🐘 PHP Configuration:"
echo "   • PHP Version: $PHP_VERSION"
echo "   • PHP CLI Config: /etc/php/$PHP_VERSION/cli/php.ini (tuned)"
echo "   • PHP-FPM Config: /etc/php/$PHP_VERSION/fpm/php.ini (tuned)"
echo "   • OPcache: Enabled with JIT"
echo "   • Memory Limit: 1024M"
echo "   • Max Execution Time: 6000s"
echo "   • Upload Max Filesize: 100M"
if [ "$IONCUBE_SKIP" != "true" ]; then
    echo "   • IonCube Loader: Enabled"
else
    echo "   • IonCube Loader: Skipped (version not available)"
fi
echo ""
echo "🔐 File Permissions:"
echo "   • Owner: www-data:www-data"
echo "   • Folders: 755 (rwxr-xr-x)"
echo "   • Files: 644 (rw-r--r--)"
echo "   • wp-config.php: 440 (r--r-----)"
echo ""

if [ "$SSL_ENABLED" = true ]; then
    echo "🔒 SSL Certificate:"
    echo "   • Path: /etc/letsencrypt/live/$DOMAIN_NAME"
    echo "   • Status: ✅ Enabled"
    echo "   • Method: Nginx"
    echo ""
    echo "🌐 Access URL: https://$DOMAIN_NAME"
else
    echo "🔓 SSL Certificate:"
    echo "   • Status: ❌ Disabled (HTTP only)"
    echo ""
    echo "🌐 Access URL: http://$DOMAIN_NAME"
fi

# Konfigurasi Tuning (PHP, MySQL, IonCube)
echo ""
echo "⚙️ Mengkonfigurasi tuning untuk performa optimal..."

# Konfigurasi PHP tuning
echo "🐘 Mengkonfigurasi PHP tuning..."
echo "📝 Mengkonfigurasi PHP CLI (/etc/php/$PHP_VERSION/cli/php.ini)..."
sudo tee /etc/php/$PHP_VERSION/cli/php.ini > /dev/null <<EOF
engine = On
short_open_tag = On
precision = 14
output_buffering = 4096
zlib.output_compression = Off
implicit_flush = Off
unserialize_callback_func =
serialize_precision = 17
disable_functions =
disable_classes =
zend.enable_gc = On
expose_php = Off
upload_max_filesize = 100M
max_execution_time = 6000
max_input_time = 6000
max_input_vars = 5000
memory_limit = 1024M
error_reporting = E_ALL & ~E_DEPRECATED & ~E_STRICT
display_errors = Off
display_startup_errors = Off
log_errors = On
log_errors_max_len = 1024
ignore_repeated_errors = Off
ignore_repeated_source = Off
report_memleaks = On
track_errors = Off
html_errors = On
variables_order = "GPCS"
request_order = "GP"
register_argc_argv = Off
auto_globals_jit = On
post_max_size = 100M
auto_prepend_file =
auto_append_file =
default_mimetype = "text/html"
default_charset = "UTF-8"
doc_root =
user_dir =
enable_dl = Off
cgi.fix_pathinfo=0
file_uploads = On
upload_tmp_dir = /tmp
max_file_uploads = 500
allow_url_fopen = On
allow_url_include = Off
default_socket_timeout = 90
date.timezone = Asia/Jakarta
pdo_mysql.cache_size = 2000
pdo_mysql.default_socket=
mysqli.max_persistent = -1
mysqli.allow_persistent = On
mysqli.max_links = -1
mysqli.cache_size = 2000
mysqli.default_port = 3306
mysqli.default_socket =
mysqli.default_host =
mysqli.default_user =
mysqli.default_pw =
mysqli.reconnect = Off
session.save_handler = files
session.use_strict_mode = 0
session.use_cookies = 1
session.use_only_cookies = 1
session.name = PHPSESSID
session.auto_start = 0
session.cookie_lifetime = 0
session.cookie_path = /
session.cookie_domain =
session.cookie_httponly =
session.serialize_handler = php
session.gc_probability = 0
session.gc_divisor = 1000
session.gc_maxlifetime = 1440
session.referer_check =
session.cache_limiter = nocache
session.cache_expire = 180
session.use_trans_sid = 0
session.hash_function = 0
session.hash_bits_per_character = 5
url_rewriter.tags = "a=href,area=href,frame=src,input=src,form=fakeentry"
; opcache.blacklist_filename = no value
opcache.consistency_checks = 0
opcache.dups_fix = Off
opcache.enable = 1
opcache.enable_cli = Off
opcache.enable_file_override = Off
opcache.error_log = /tmp/opcache/error-opcache.log
; opcache.file_cache = no value
opcache.file_cache_consistency_checks = On
opcache.file_cache_only = Off
opcache.file_update_protection = 2
opcache.force_restart_timeout = 180
opcache.huge_code_pages = Off
opcache.interned_strings_buffer = 8
opcache.jit = tracing
opcache.jit_bisect_limit = 0
opcache.jit_blacklist_root_trace = 16
opcache.jit_blacklist_side_trace = 8
opcache.jit_buffer_size = 100M
opcache.jit_debug = 0
opcache.jit_hot_func = 127
opcache.jit_hot_loop = 64
opcache.jit_hot_return = 8
opcache.jit_hot_side_exit = 8
opcache.jit_max_exit_counters = 8192
opcache.jit_max_loop_unrolls = 8
opcache.jit_max_polymorphic_calls = 2
opcache.jit_max_recursive_calls = 2
opcache.jit_max_recursive_returns = 2
opcache.jit_max_root_traces = 1024
opcache.jit_max_side_traces = 128
opcache.jit_prof_threshold = 0.005
opcache.lockfile_path = /tmp
opcache.log_verbosity_level = 1
opcache.max_accelerated_files = 10000
opcache.max_file_size = 0
opcache.max_wasted_percentage = 5
opcache.memory_consumption = 128
opcache.opt_debug_level = 0
opcache.optimization_level = 0x7FFEBFFF
; opcache.preferred_memory_model = no value
; opcache.preload = no value
; opcache.preload_user = no value
opcache.protect_memory = Off
opcache.record_warnings = Off
; opcache.restrict_api = no value
opcache.revalidate_freq = 60
opcache.revalidate_path = Off
opcache.save_comments = On
opcache.use_cwd = On
opcache.validate_permission = Off
opcache.validate_root = Off
opcache.validate_timestamps = On
EOF

echo "📝 Mengkonfigurasi PHP-FPM (/etc/php/$PHP_VERSION/fpm/php.ini)..."
sudo tee /etc/php/$PHP_VERSION/fpm/php.ini > /dev/null <<EOF
engine = On
short_open_tag = On
precision = 14
output_buffering = 4096
zlib.output_compression = Off
implicit_flush = Off
unserialize_callback_func =
serialize_precision = 17
disable_functions =
disable_classes =
zend.enable_gc = On
expose_php = Off
upload_max_filesize = 100M
max_execution_time = 6000
max_input_time = 6000
max_input_vars = 5000
memory_limit = 1024M
error_reporting = E_ALL & ~E_DEPRECATED & ~E_STRICT
display_errors = Off
display_startup_errors = Off
log_errors = On
log_errors_max_len = 1024
ignore_repeated_errors = Off
ignore_repeated_source = Off
report_memleaks = On
track_errors = Off
html_errors = On
variables_order = "GPCS"
request_order = "GP"
register_argc_argv = Off
auto_globals_jit = On
post_max_size = 100M
auto_prepend_file =
auto_append_file =
default_mimetype = "text/html"
default_charset = "UTF-8"
doc_root =
user_dir =
enable_dl = Off
cgi.fix_pathinfo=0
file_uploads = On
upload_tmp_dir = /tmp
max_file_uploads = 500
allow_url_fopen = On
allow_url_include = Off
default_socket_timeout = 90
date.timezone = Asia/Jakarta
pdo_mysql.cache_size = 2000
pdo_mysql.default_socket=
mysqli.max_persistent = -1
mysqli.allow_persistent = On
mysqli.max_links = -1
mysqli.cache_size = 2000
mysqli.default_port = 3306
mysqli.default_socket =
mysqli.default_host =
mysqli.default_user =
mysqli.default_pw =
mysqli.reconnect = Off
session.save_handler = files
session.use_strict_mode = 0
session.use_cookies = 1
session.use_only_cookies = 1
session.name = PHPSESSID
session.auto_start = 0
session.cookie_lifetime = 0
session.cookie_path = /
session.cookie_domain =
session.cookie_httponly =
session.serialize_handler = php
session.gc_probability = 0
session.gc_divisor = 1000
session.gc_maxlifetime = 1440
session.referer_check =
session.cache_limiter = nocache
session.cache_expire = 180
session.use_trans_sid = 0
session.hash_function = 0
session.hash_bits_per_character = 5
url_rewriter.tags = "a=href,area=href,frame=src,input=src,form=fakeentry"
; opcache.blacklist_filename = no value
opcache.consistency_checks = 0
opcache.dups_fix = Off
opcache.enable = 1
opcache.enable_cli = Off
opcache.enable_file_override = Off
opcache.error_log = /tmp/opcache/error-opcache.log
; opcache.file_cache = no value
opcache.file_cache_consistency_checks = On
opcache.file_cache_only = Off
opcache.file_update_protection = 2
opcache.force_restart_timeout = 180
opcache.huge_code_pages = Off
opcache.interned_strings_buffer = 8
opcache.jit = tracing
opcache.jit_bisect_limit = 0
opcache.jit_blacklist_root_trace = 16
opcache.jit_blacklist_side_trace = 8
opcache.jit_buffer_size = 100M
opcache.jit_debug = 0
opcache.jit_hot_func = 127
opcache.jit_hot_loop = 64
opcache.jit_hot_return = 8
opcache.jit_hot_side_exit = 8
opcache.jit_max_exit_counters = 8192
opcache.jit_max_loop_unrolls = 8
opcache.jit_max_polymorphic_calls = 2
opcache.jit_max_recursive_calls = 2
opcache.jit_max_recursive_returns = 2
opcache.jit_max_root_traces = 1024
opcache.jit_max_side_traces = 128
opcache.jit_prof_threshold = 0.005
opcache.lockfile_path = /tmp
opcache.log_verbosity_level = 1
opcache.max_accelerated_files = 10000
opcache.max_file_size = 0
opcache.max_wasted_percentage = 5
opcache.memory_consumption = 128
opcache.opt_debug_level = 0
opcache.optimization_level = 0x7FFEBFFF
; opcache.preferred_memory_model = no value
; opcache.preload = no value
; opcache.preload_user = no value
opcache.protect_memory = Off
opcache.record_warnings = Off
; opcache.restrict_api = no value
opcache.revalidate_freq = 60
opcache.revalidate_path = Off
opcache.save_comments = On
opcache.use_cwd = On
opcache.validate_permission = Off
opcache.validate_root = Off
opcache.validate_timestamps = On
EOF

# Buat direktori untuk opcache error log
echo "📁 Membuat direktori untuk opcache error log..."
sudo mkdir -p /tmp/opcache
sudo chown www-data:www-data /tmp/opcache
sudo chmod 755 /tmp/opcache

echo "✅ PHP tuning configuration applied"


# Install IonCube Loader (jika dipilih)
if [ "$IONCUBE_ENABLED" = true ]; then
    echo "🧩 Installing IonCube Loader..."
    
    # Inisialisasi variabel IonCube
    IONCUBE_SKIP=false

# Deteksi extension directory
echo "🔍 Mendeteksi PHP extension directory..."
PHP_EXT_DIR=$(php -i | grep "extension_dir" | awk '{print $3}' | head -1)
echo "📁 PHP Extension Directory: $PHP_EXT_DIR"

# Download IonCube Loader
echo "⬇️ Mengunduh IonCube Loader..."
cd /tmp
if [ ! -f "ioncube_loaders_lin_x86-64.tar.gz" ]; then
    sudo wget https://downloads.ioncube.com/loader_downloads/ioncube_loaders_lin_x86-64.tar.gz
    echo "✅ IonCube Loader berhasil diunduh"
else
    echo "✅ IonCube Loader sudah tersedia"
fi

# Extract IonCube Loader
echo "📦 Mengekstrak IonCube Loader..."
sudo tar -xzf ioncube_loaders_lin_x86-64.tar.gz

# Copy IonCube loader untuk versi PHP yang sesuai
echo "📋 Menyalin IonCube loader untuk PHP $PHP_VERSION..."
IONCUBE_LOADER="ioncube_loader_lin_${PHP_VERSION}.so"
if [ -f "ioncube/$IONCUBE_LOADER" ]; then
    sudo cp "ioncube/$IONCUBE_LOADER" "$PHP_EXT_DIR/"
    echo "✅ IonCube loader berhasil disalin: $IONCUBE_LOADER"
else
    echo "⚠️  IonCube loader untuk PHP $PHP_VERSION tidak ditemukan"
    echo "🔍 Mencari loader yang tersedia..."
    ls -la ioncube/ioncube_loader_lin_*.so 2>/dev/null || echo "❌ Tidak ada loader yang tersedia"
    echo "💡 IonCube loader akan dilewati"
    IONCUBE_SKIP=true
fi

# Konfigurasi IonCube di PHP CLI
if [ "$IONCUBE_SKIP" != "true" ]; then
    echo "📝 Mengkonfigurasi IonCube di PHP CLI..."
    if ! grep -q "ioncube_loader" "/etc/php/$PHP_VERSION/cli/php.ini"; then
        sudo sed -i "1i zend_extension = \"$PHP_EXT_DIR/$IONCUBE_LOADER\"" "/etc/php/$PHP_VERSION/cli/php.ini"
        echo "✅ IonCube loader ditambahkan ke PHP CLI config"
    else
        echo "ℹ️  IonCube loader sudah ada di PHP CLI config"
    fi
fi

# Konfigurasi IonCube di PHP-FPM
if [ "$IONCUBE_SKIP" != "true" ]; then
    echo "📝 Mengkonfigurasi IonCube di PHP-FPM..."
    if ! grep -q "ioncube_loader" "/etc/php/$PHP_VERSION/fpm/php.ini"; then
        sudo sed -i "1i zend_extension = \"$PHP_EXT_DIR/$IONCUBE_LOADER\"" "/etc/php/$PHP_VERSION/fpm/php.ini"
        echo "✅ IonCube loader ditambahkan ke PHP-FPM config"
    else
        echo "ℹ️  IonCube loader sudah ada di PHP-FPM config"
    fi
fi
    if [ "$IONCUBE_SKIP" != "true" ]; then
        echo "✅ IonCube Loader installation completed"
    else
        echo "⚠️  IonCube Loader installation skipped"
    fi
else
    echo "ℹ️  IonCube Loader dilewati (tidak dipilih)"
    IONCUBE_SKIP=true
fi

# Tuning tambahan PHP-FPM pool dan OPCache
echo ""
echo "🧰 Mengaplikasikan tuning tambahan PHP-FPM & OPCache..."
PHP_FPM_POOL="/etc/php/$PHP_VERSION/fpm/pool.d/www.conf"
OPCACHE_INI="/etc/php/$PHP_VERSION/fpm/conf.d/10-opcache.ini"

echo "   • Menyisipkan konfigurasi pool di $PHP_FPM_POOL"
sudo sed -i "/^;* Custom FPM Tuning START/,/^;* Custom FPM Tuning END/d" "$PHP_FPM_POOL"
sudo tee -a "$PHP_FPM_POOL" > /dev/null <<EOF
; Custom FPM Tuning START
pm = dynamic
pm.max_children = 80
pm.start_servers = 10
pm.min_spare_servers = 10
pm.max_spare_servers = 20
pm.max_requests = 500
; Custom FPM Tuning END
EOF

echo "   • Mengonfigurasi OPCache di $OPCACHE_INI"
sudo mkdir -p "$(dirname "$OPCACHE_INI")"
sudo touch "$OPCACHE_INI"
sudo sed -i "/^;* Custom OPCache Tuning START/,/^;* Custom OPCache Tuning END/d" "$OPCACHE_INI"
sudo tee -a "$OPCACHE_INI" > /dev/null <<EOF
; Custom OPCache Tuning START
opcache.enable=1
opcache.memory_consumption=1024
opcache.interned_strings_buffer=128
opcache.max_accelerated_files=50000
opcache.validate_timestamps=0
; Custom OPCache Tuning END
EOF

# Restart services untuk menerapkan semua konfigurasi
echo "🔄 Merestart services untuk menerapkan konfigurasi tuning..."
sudo systemctl restart php${PHP_VERSION}-fpm

# Restart semua versi PHP-FPM yang tersedia (jika ada)
if compgen -G "/lib/systemd/system/php*-fpm.service" > /dev/null; then
    for unit in /lib/systemd/system/php*-fpm.service; do
        SERVICE_NAME=$(basename "$unit" .service)
        sudo systemctl restart "$SERVICE_NAME" || true
    done
fi

sudo systemctl restart nginx

echo "✅ All tuning configurations applied successfully!"

echo "🎉 Setup completed successfully!"