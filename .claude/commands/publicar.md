---
allowed-tools: Bash(npm install:*), Bash(npm run build:*), Bash(git add:*), Bash(git commit:*), Bash(git push:*), Bash(git status:*), Bash(git log:*), Bash(git diff:*)
description: Reemplaza App.jsx con contaduria.jsx, instala deps, compila, commitea y pushea
model: claude-haiku-4-5-20251001
---

Actualicé el archivo contaduria.jsx en esta carpeta con cambios nuevos.

Por favor:
1. Reemplazá el contenido de src/App.jsx por el de contaduria.jsx.
2. Corré npm install (por si hay alguna dependencia nueva en package.json
   que todavía no esté instalada).
3. Corré npm run build y confirmá que compila sin errores.
4. Si compila bien, hacé git add, commit (mensaje corto describiendo el
   cambio) y push a master.
5. Avisame cuando el push haya terminado para que yo chequee el redeploy
   en Vercel.