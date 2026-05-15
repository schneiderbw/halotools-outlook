# syntax=docker/dockerfile:1.7
# Build stage — build @iusehalo/halo-api once, then both Vite apps off it.
FROM node:20-alpine AS build
WORKDIR /app

# Workspace skeleton first so `npm ci` resolves cross-workspace deps without
# pulling in source. Each app's package.json gets copied into its workspace
# directory, then `npm ci` reads the root lockfile to install everything.
COPY package.json package-lock.json* ./
COPY packages/halo-api/package.json packages/halo-api/package.json
COPY apps/outlook/package.json apps/outlook/package.json
COPY apps/teams/package.json apps/teams/package.json

RUN npm ci || npm install

# Now the rest of the source — keeps the dep install layer cacheable.
COPY . .

# Builds in order: halo-api → outlook → teams.
RUN npm run build:outlook && npm run build:teams

# Runtime stage — single nginx serving all three surfaces:
#   /           → marketing landing page (plain HTML)
#   /outlook/   → Outlook add-in SPA
#   /teams/     → Teams app SPA
FROM nginx:1.27-alpine
COPY --from=build /app/apps/outlook/dist /usr/share/nginx/html/outlook
COPY --from=build /app/apps/teams/dist /usr/share/nginx/html/teams
COPY landing/index.html /usr/share/nginx/html/index.html
COPY nginx.conf /etc/nginx/conf.d/default.conf
EXPOSE 80
