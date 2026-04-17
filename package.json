FROM node:20-slim

WORKDIR /app

# Install root dependencies
COPY package.json ./
RUN npm install --production=false && echo "rebuild-1"

# Install client dependencies and build
COPY client/package.json client/
RUN cd client && npm install && echo "rebuild-1"

COPY client/ client/
RUN cd client && npx vite build

# Copy server
COPY server/ server/
COPY .env.example .env

# Create data directory for SQLite and attachments
RUN mkdir -p data/attachments data/entity_files

ENV NODE_ENV=production
ENV PORT=3000

EXPOSE 3000

CMD ["node", "server/index.js"]
