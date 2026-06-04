# GitHub CLI cheat sheet

Handy [`gh`](https://cli.github.com/) commands for working with GitHub from the terminal.

## Auth

| Command | What it does |
| --- | --- |
| `gh auth login` | Log in to a GitHub account |
| `gh auth status` | Show the current login state |

## Repositories

| Command | What it does |
| --- | --- |
| `gh repo list` | List your repositories |
| `gh repo create <name> --public --source=. --push` | Create a repo from the current folder |
| `gh repo clone <owner>/<repo>` | Clone a repository |

## Pull requests

| Command | What it does |
| --- | --- |
| `gh pr create` | Open a pull request |
| `gh pr checks <n> --watch` | Watch a PR's CI checks |
| `gh pr merge <n> --squash --delete-branch` | Squash-merge a PR and delete its branch |

## See also

- [Git cheat sheet](git-cheatsheet.md)
- [Official gh manual](https://cli.github.com/manual/)
