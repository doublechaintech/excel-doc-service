git filter-branch --force --index-filter 'git rm --cached --ignore-unmatch excel-doc-service' --prune-empty --tag-name-filter cat -- --all
echo "excel-doc-service" >> .gitignore
git add  .gitignore
git commit -m "remove excel-doc-service"
git push origin --force --all
