#!/bin/bash
# Lanzador para subir a GitHub automáticamente
cd "$(dirname "$0")"
echo "Subiendo proyecto a GitHub..."
git push -u origin main
read -p "Presiona Enter para cerrar esta ventana..."
