# ==============================================================================
# React Frontend Dockerfile (Multi-Stage Build - Step 3 & Supplementary 2)
# ==============================================================================

# ------------------------------------------------------------------------------
# Stage 1: Build Stage (Node 20+ required for Vite / modern node:util styleText)
# ------------------------------------------------------------------------------
FROM node:20-alpine AS builder

WORKDIR /app

# Copy package manifests and install full dependencies (including Vite)
COPY package*.json ./
RUN npm install

# Copy application source code and compile production assets
COPY . .
RUN npm run build

# ------------------------------------------------------------------------------
# Stage 2: Lean Production Runtime Stage
# ------------------------------------------------------------------------------
FROM node:20-alpine

WORKDIR /app

# Copy package manifests
COPY package*.json ./

# Copy compiled dist artifacts and runtime modules from builder stage
COPY --from=builder /app/node_modules ./node_modules
COPY --from=builder /app/dist ./dist

# Expose frontend server port
EXPOSE 5173

# Serve production bundle on port 5173 accessible to external network
CMD ["npm", "run", "preview", "--", "--host", "0.0.0.0", "--port", "5173"]
