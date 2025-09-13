FROM php:8.2-apache

# Enable .htaccess and mod_rewrite
RUN a2enmod rewrite

# Copy application files
COPY ./index.html /var/www/html/
COPY ./.htaccess /var/www/html/
COPY ./privacy.html /var/www/html/
COPY ./terms.html /var/www/html/
COPY ./testing.html /var/www/html/
COPY ./bgvideos-leaveform.mp4 /var/www/html/

# Copy folders
COPY ./down /var/www/html/down
COPY ./.well-known /var/www/html/.well-known
COPY ./images /var/www/html/images

# Allow .htaccess overrides globally
RUN sed -i '/<Directory \/var\/www\/>/,/<\/Directory>/ s/AllowOverride None/AllowOverride All/' /etc/apache2/apache2.conf

# Set secure permissions
RUN chown -R www-data:www-data /var/www/html/ && chmod -R 755 /var/www/html/