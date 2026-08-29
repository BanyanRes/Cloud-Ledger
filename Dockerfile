FROM node:20-slim

# OCR tooling for image-only bank statement PDFs (scanned copies, or bank
# portals like MapleMark that export statements as page images with no text
# layer): poppler-utils provides pdftoppm to rasterize pages, tesseract-ocr
# reads the text back out.
RUN apt-get update \
 && apt-get install -y --no-install-recommends tesseract-ocr poppler-utils \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install root dependencies
COPY package.json ./
RUN npm install --production=false && echo "rebuild-2"

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
