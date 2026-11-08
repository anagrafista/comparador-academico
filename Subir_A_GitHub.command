#!/bin/bash
# Lanzador para subir a GitHub automáticamente
cd "$(dirname "$0")"
echo "Sincronizando con GitHub..."
git pull --rebase origin main
echo "Subiendo cambios..."
git push origin main
read -p "Presiona Enter para cerrar esta ventana..."
