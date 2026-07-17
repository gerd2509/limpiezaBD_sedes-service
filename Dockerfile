FROM node:20-alpine
WORKDIR /app

# Instala dependencias primero (mejor caché de capas)
COPY package*.json ./
RUN npm install --omit=dev

# Copia el código
COPY . .

EXPOSE 4002
ENV PORT=4002
CMD ["node", "index.js"]
