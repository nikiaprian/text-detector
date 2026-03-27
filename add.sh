#!/bin/bash

# Exit jika error dan gunakan variabel dengan aman
set -euo pipefail

echo "🌐 Add WordPress Domain - Nginx Plugin Method (Auto-Configure SSL)"
echo "===================================================================="
echo ""

# Input domain
read -p "Masukkan nama domain (contoh: example.com): " DOMAIN_NAME

# Generate random password MySQL (20 karakter)
MYSQL_PASSWORD=$(openssl rand -base64 32 | tr -d "=+/" | cut -c1-20)

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

# Cek apakah domain sudah ada
if [ -d "/var/www/html/$DOMAIN_NAME" ]; then
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

# Sanitize domain name for database identifiers (remove dots and hyphens)
DOMAIN_SANITIZED=$(echo "$DOMAIN_NAME" | tr -d '.-')
DB_NAME="db_${DOMAIN_SANITIZED}"
# Generate DB_USER with maximum 32 characters (MySQL limit)
# Use first 8 chars + hash of domain to ensure uniqueness while staying under limit
DOMAIN_CLEAN="${DOMAIN_SANITIZED}"
DOMAIN_HASH=$(echo -n "$DOMAIN_NAME" | md5sum | cut -c1-8)
DB_USER="u_$(echo "$DOMAIN_CLEAN" | cut -c1-20)_$DOMAIN_HASH"
# Ensure DB_USER doesn't exceed 32 characters
if [ ${#DB_USER} -gt 32 ]; then
    DB_USER="u_$(echo "$DOMAIN_CLEAN" | cut -c1-15)_$DOMAIN_HASH"
fi
REDIS_PREFIX=$(echo "$DOMAIN_SANITIZED" | cut -c1-8)
WP_CONFIG="/var/www/html/$DOMAIN_NAME/wp-config-sample.php"

# Install Nginx jika belum ada
if ! command -v nginx &> /dev/null; then
  echo "⚙️ Installing Nginx..."
  sudo apt update
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

# setup db
echo "🧩 Setup database dan user..."
sudo mysql <<EOF
DROP DATABASE IF EXISTS $DB_NAME;
DROP USER IF EXISTS '$DB_USER'@'localhost';
CREATE DATABASE $DB_NAME DEFAULT CHARACTER SET utf8 COLLATE utf8_unicode_ci;
CREATE USER '$DB_USER'@'localhost' IDENTIFIED BY '$MYSQL_PASSWORD';
GRANT ALL ON $DB_NAME.* TO '$DB_USER'@'localhost';
FLUSH PRIVILEGES;
EOF

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

# Konfigurasi Firewall
echo "🛡️ Mengatur firewall..."
sudo ufw allow 'Nginx Full'

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

# Setup Nginx configuration (HTTP first)
echo ""
echo "📝 Membuat konfigurasi Nginx HTTP untuk $DOMAIN_NAME..."

# Buat direktori sites-available dan sites-enabled jika belum ada
sudo mkdir -p /etc/nginx/sites-available
sudo mkdir -p /etc/nginx/sites-enabled

# Hapus konfigurasi lama jika ada
if [ -f "/etc/nginx/sites-available/$DOMAIN_NAME" ]; then
    echo "⚠️  Konfigurasi Nginx untuk '$DOMAIN_NAME' sudah ada. Backup & hapus konfigurasi lama..."
    sudo cp "/etc/nginx/sites-available/$DOMAIN_NAME" "/etc/nginx/sites-available/$DOMAIN_NAME.old.$(date +%Y%m%d_%H%M%S)"
    sudo rm -f "/etc/nginx/sites-available/$DOMAIN_NAME"
    sudo rm -f "/etc/nginx/sites-enabled/$DOMAIN_NAME"
fi

# Buat konfigurasi HTTP-only (Certbot akan upgrade ke HTTPS)
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

    # Security: Block access to hidden files
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

# Test konfigurasi Nginx
echo "🧪 Testing Nginx configuration..."
if sudo nginx -t; then
    echo "✅ Nginx configuration valid"
else
    echo "❌ Nginx configuration error"
    exit 1
fi

# Start/Reload Nginx dan PHP-FPM
echo "🔄 Reload Nginx dan PHP-FPM..."
sudo systemctl reload php${PHP_VERSION}-fpm 2>/dev/null || sudo systemctl start php${PHP_VERSION}-fpm
sudo systemctl reload nginx 2>/dev/null || sudo systemctl start nginx

# Enable services untuk auto-start
sudo systemctl enable nginx
sudo systemctl enable php${PHP_VERSION}-fpm

echo "✅ Nginx dan PHP-FPM sudah running"

# Setup SSL dengan Certbot Nginx Plugin jika enabled
if [ "$SSL_ENABLED" = true ]; then
    echo ""
    echo "🔐 Mengaktifkan SSL dengan Certbot Nginx Plugin..."
    echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    echo "ℹ️  Method: Certbot --nginx (Auto-configure HTTPS)"
    
    if [ ! -d "/etc/letsencrypt/live/$DOMAIN_NAME" ]; then
        echo "🔐 Mendapatkan SSL certificate untuk $DOMAIN_NAME..."
        
        # Gunakan certbot --nginx untuk auto-configure HTTPS
        if sudo certbot --nginx \
            --non-interactive \
            --agree-tos \
            --email "$EMAIL" \
            -d "$DOMAIN_NAME" \
            --redirect; then
            
            echo "✅ SSL certificate berhasil diperoleh dan dikonfigurasi!"
            echo "✅ HTTPS telah diaktifkan dengan auto-redirect dari HTTP"
            echo "ℹ️  Certbot telah mengupdate Nginx config secara otomatis"
            
            # Tambahkan SSL optimization dan security headers
            echo "🔒 Menambahkan SSL optimization dan security headers..."
            
            # Backup config yang dibuat certbot
            sudo cp "/etc/nginx/sites-available/$DOMAIN_NAME" "/etc/nginx/sites-available/$DOMAIN_NAME.certbot-backup"
            
            # Tambahkan SSL optimization dan security headers
            if ! grep -q "Strict-Transport-Security" "/etc/nginx/sites-available/$DOMAIN_NAME"; then
                # Sisipkan header keamanan setelah baris ssl_certificate_key; opsi SSL lain sudah diatur oleh Certbot
                sudo sed -i "/ssl_certificate_key/a \\\n    # Security Headers\n    add_header Strict-Transport-Security \"max-age=31536000; includeSubDomains\" always;\n    add_header X-Frame-Options \"SAMEORIGIN\" always;\n    add_header X-Content-Type-Options \"nosniff\" always;\n    add_header X-XSS-Protection \"1; mode=block\" always;" "/etc/nginx/sites-available/$DOMAIN_NAME"
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
                echo "💡 Recommendation: Review config manually or use add-webroot.sh for full control"
                echo "📁 Config file: /etc/nginx/sites-available/$DOMAIN_NAME"
                echo "📋 Backup HTTP config: /etc/nginx/sites-available/$DOMAIN_NAME.old.*"
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

# Cek status service
echo ""
echo "📊 Status Nginx service:"
sudo systemctl status nginx --no-pager -l

echo ""
echo "📊 Status PHP-FPM service:"
sudo systemctl status php${PHP_VERSION}-fpm --no-pager -l

# Info akhir
echo ""
echo "✅ Instalasi berhasil!"
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
echo ""
echo "🔐 File Permissions:"
echo "   • Owner: www-data:www-data"
echo "   • Folders: 755 (rwxr-xr-x)"
echo "   • Files: 644 (rw-r--r--)"
echo "   • wp-config.php: 440 (r--r-----)"
echo ""

if [ "$SSL_ENABLED" = true ]; then
    if [ -d "/etc/letsencrypt/live/$DOMAIN_NAME" ]; then
        echo "🔒 SSL Certificate:"
        echo "   • Path: /etc/letsencrypt/live/$DOMAIN_NAME"
        echo "   • Status: ✅ Enabled"
        echo "   • Method: Nginx"
        echo "🌐 Access URL: https://$DOMAIN_NAME"
    else
        echo "🔓 SSL Certificate:"
        echo "   • Status: ❌ Failed to obtain"
        echo "🌐 Access URL: http://$DOMAIN_NAME"
    fi
else
    echo "🔓 SSL Certificate:"
    echo "   • Status: ❌ Disabled (HTTP only)"
    echo "🌐 Access URL: http://$DOMAIN_NAME"
fi

if [ "$SSL_ENABLED" = true ]; then
    echo ""
    echo "🔐 SSL Management Commands:"
    echo "   • Check SSL status: sudo certbot certificates"
    echo ""
    echo "⚠️  Note: Certbot --nginx auto-modified your Nginx config"
    echo "   • Backup tersimpan di: /etc/nginx/sites-available/$DOMAIN_NAME.certbot-backup"
fi

echo ""
echo "🎉 Setup completed successfully!"