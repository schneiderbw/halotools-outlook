# syntax=docker/dockerfile:1.7
# Build stage
FROM node:20-alpine AS build
WORKDIR /app
COPY package.json package-lock.json* ./
RUN npm ci || npm install
COPY . .
RUN npm run build

# Runtime stage — nginx serving /outlook/ for the add-in and / for the landing page
FROM nginx:1.27-alpine
COPY --from=build /app/dist /usr/share/nginx/html/outlook
# Marketing / landing page at the root. Plain HTML, no build step required.
COPY landing/index.html /usr/share/nginx/html/index.html
COPY nginx.conf /etc/nginx/conf.d/default.conf
EXPOSE 80
