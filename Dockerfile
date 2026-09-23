FROM node:20-bookworm-slim

# LibreOffice (for QA rendering + validation) and Python/python-pptx (for extraction)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    python3 \
    python3-pip \
    poppler-utils \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

RUN pip3 install --break-system-packages --no-cache-dir python-pptx

WORKDIR /app
COPY package.json package-lock.json* ./
RUN npm install --omit=dev

COPY . .

ENV PORT=3000
EXPOSE 3000

CMD ["node", "server.js"]
