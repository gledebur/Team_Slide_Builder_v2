#!/bin/bash

echo "🚀 Starting Team Slide Generator (Frontend + Backend)..."

# Function to cleanup background processes
cleanup() {
    echo "🛑 Shutting down services..."
    # Kill all background jobs
    jobs -p | xargs -r kill
    exit 0
}

# Set up signal handlers
trap cleanup SIGINT SIGTERM

# Start Backend
echo "📊 Starting Flask Backend..."
cd backend

# Check if virtual environment exists
if [ ! -d "venv" ]; then
    echo "Creating Python virtual environment..."
    python3 -m venv venv
fi

# Activate virtual environment
echo "Activating virtual environment..."
source venv/bin/activate

# Install dependencies
echo "Installing Python dependencies..."
pip install -r requirements.txt

# Start Flask app in background
echo "Starting Flask server on http://localhost:5000"
python app.py &
BACKEND_PID=$!

# Wait a moment for backend to start
sleep 3

# Go back to root directory
cd ..

# Start Frontend
echo "🌐 Starting Next.js Frontend..."

# Check if node_modules exists
if [ ! -d "node_modules" ]; then
    echo "Installing Node.js dependencies..."
    npm install
fi

# Start Next.js app
echo "Starting Next.js development server on http://localhost:3000"
npm run dev &
FRONTEND_PID=$!

echo ""
echo "✅ Services started successfully!"
echo "🌐 Frontend: http://localhost:3000"
echo "📊 Backend: http://localhost:5000"
echo ""
echo "Press Ctrl+C to stop all services"

# Wait for background processes
wait