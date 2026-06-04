# Git cheat sheet

A few git commands I reach for often.

## Everyday

| Command | What it does |
| --- | --- |
| `git status` | Show the working tree status |
| `git add <file>` | Stage a file |
| `git commit -m "msg"` | Commit staged changes |
| `git push` | Push commits to the remote |

## Branches

| Command | What it does |
| --- | --- |
| `git switch -c <name>` | Create and switch to a new branch |
| `git switch <name>` | Switch to an existing branch |
| `git merge <name>` | Merge a branch into the current one |

## Undo

| Command | What it does |
| --- | --- |
| `git restore <file>` | Discard changes in a file |
| `git reset --soft HEAD~1` | Undo the last commit, keep changes staged |
