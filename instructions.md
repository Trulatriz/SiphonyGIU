# Instructions for committing and pushing changes

## 1. Configure Git if needed
- git config --global user.name "Your Name"
- git config --global user.email "your.email@example.com"

## 2. Initialize the repository (first time only)
- Run git rev-parse --show-toplevel to check if Git is already initialized.
- If it reports an error, run git init in the project root.

## 3. Record the GitHub token securely
- Set an environment variable (PowerShell example):
  `powershell
   = "<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>"
  `
- Or create/update a credential helper (git config credential.helper manager) so you can paste the token once when prompted.

## 4. Add the GitHub remote
- Remove any existing origin if it points elsewhere: git remote remove origin (ignore errors if it doesn’t exist).
- Add the new remote using the token placeholder:
  `powershell
  git remote add origin https://@github.com/Trulatriz/SiphonyGIU.git
  `
- Alternatively, use the standard URL and enter the token when Git prompts for a password:
  `powershell
  git remote add origin https://github.com/Trulatriz/SiphonyGIU.git
  `

## 5. Review changes
- git status
- git diff

## 6. Stage and commit
- Stage selectively or everything: git add path/to/file or git add .
- Commit with a descriptive message: git commit -m "Describe your change"

## 7. Push to GitHub
- First push creates the main branch upstream:
  `powershell
  git push -u origin main
  `
- Subsequent pushes can omit the -u flag: git push

## 8. If authentication fails
- Ensure the $env:GITHUB_TOKEN variable is still set in the current session.
- Tokens act as passwords; use the provided one as-is when prompted.
- Tokens can be revoked or regenerated in GitHub settings if needed.

## 9. Keep credentials safe
- Do not commit the token or place it in tracked files.
- Clear the environment variable when finished: Remove-Item Env:GITHUB_TOKEN
- Consider using a .env file in .gitignore if you want to store the token locally for scripts.
