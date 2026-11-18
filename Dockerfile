FROM node:22-alpine

# Zona waktu aplikasi = zona waktu script Apps Script lama (Asia/Jakarta).
RUN apk add --no-cache tzdata
ENV TZ=Asia/Jakarta \
    NODE_ENV=production

WORKDIR /app
COPY package.json package-lock.json* ./
RUN npm ci --omit=dev || npm install --omit=dev

COPY server ./server
COPY scripts ./scripts
COPY public ./public

USER node
EXPOSE 3000
CMD ["node", "server/index.js"]
