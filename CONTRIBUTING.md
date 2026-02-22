# Contributing to DeepExtract

Thanks for your interest in contributing.

## How to contribute

- Fork this repository
- Create a feature branch
- Make focused changes with clear commit messages
- Open a Pull Request with a short description

## Local setup

```bash
pip install -r backend/requirements.txt
python start.py
```

Open `http://localhost:5000` and verify basic upload + convert flow.

## Coding guidelines

- Keep changes small and modular
- Follow existing code style in each file
- Prefer readable names over abbreviations
- Do not add hardcoded secrets

## Security and secrets

- Never commit API keys, tokens, or credentials
- Use `MINERU_API_KEY` via environment variables
- `apikey.md` is for local development only

## Pull Request checklist

- [ ] Code runs locally
- [ ] No sensitive information added
- [ ] README/docs updated if behavior changed
- [ ] PR description explains what and why
