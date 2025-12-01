#!/bin/bash

# ============================================
# Setup Script for Klader Project
# ============================================

echo "=== Klader Project Setup ==="
echo ""

# Check if git is installed
if ! command -v git &> /dev/null; then
    echo "Git is not installed. Please install Git first."
    exit 1
fi

echo "[1/4] Checking project structure..."
# Create necessary directories if they don't exist
mkdir -p Anteckningar
mkdir -p Statistik

echo "[2/4] Initializing Git repository..."
# Initialize git if not already initialized
if [ ! -d ".git" ]; then
    git init
    echo "Git repository initialized."
else
    echo "Git repository already exists."
fi

echo "[3/4] Creating .gitignore..."
# Create .gitignore if it doesn't exist
if [ ! -f ".gitignore" ]; then
    cat > .gitignore << 'EOF'
# OS generated files
.DS_Store
Thumbs.db

# IDE files
.vscode/
.idea/

# Temporary files
*.tmp
*.log
*.bak

# Environment files
.env
.env.local
EOF
    echo ".gitignore created."
else
    echo ".gitignore already exists."
fi

echo "[4/4] Creating initial commit..."
# Add all files and create initial commit
git add .
git commit -m "Initial commit: Klader project setup" 2>/dev/null || echo "No changes to commit or already committed."

echo ""
echo "=== Setup Complete ==="
echo ""
echo "Next steps to connect to GitHub:"
echo "1. Create a new repository on GitHub named 'Klader'"
echo "2. Run: git remote add origin https://github.com/Sytematic1036/Klader.git"
echo "3. Run: git branch -M main"
echo "4. Run: git push -u origin main"
echo ""
