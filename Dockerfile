FROM php:8.4-alpine

ENV COMPOSER_HOME=/composer

# 安装zip
RUN apk add --no-cache libzip-dev && \
    docker-php-ext-install -j$(nproc) zip

# 安装 GD 依赖
RUN apk add --no-cache freetype-dev libjpeg-turbo-dev libpng-dev libwebp-dev zlib-dev && \
    docker-php-ext-configure gd --with-freetype --with-jpeg --with-webp && \
    docker-php-ext-install -j$(nproc) gd

# 安装 Composer
RUN curl https://getcomposer.org/download/2.9.5/composer.phar -o /usr/local/bin/composer && \
    chmod +x /usr/local/bin/composer

WORKDIR /app
