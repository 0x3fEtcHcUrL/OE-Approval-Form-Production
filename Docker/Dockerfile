FROM php:8.2-apache

# Enable .htaccess and mod_rewrite
RUN a2enmod rewrite

# Copy your app
COPY ./index.html /var/www/html/
COPY ./.htaccess /var/www/html/
COPY ./privacy.html /var/www/html/
COPY ./terms.html /var/www/html/
COPY ./maintenance.html /var/www/html/
COPY ./bgvideos-leaveform.mp4 /var/www/html

# Allow .htaccess overrides
RUN sed -i '/<Directory \/var\/www\/>/,/<\/Directory>/ s/AllowOverride None/AllowOverride All/' /etc/apache2/apache2.conf

# Set permissions
RUN chown -R www-data:www-data /var/www/html/ && chmod -R 755 /var/www/html/