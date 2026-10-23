# ==============================================================================
# Node/Express Backend Dockerfile (Practical 11 - Step 2)
# ==============================================================================
FROM node:20-alpine

# Set working directory inside container
WORKDIR /app

# Copy package specifications first to leverage Docker layer caching
COPY package*.json ./

# Install application dependencies
RUN npm install

# Copy application source code
COPY . .

# Expose backend API port
EXPOSE 5000

# Set default environment variables
ENV PORT=5000
ENV NODE_ENV=production

# Start Express REST server
CMD ["node", "server.js"]
