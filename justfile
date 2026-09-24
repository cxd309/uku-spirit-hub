# List available recipes
default:
    @just --list

# Install the pinned dev tools (TypeScript, Apps Script types, dprint, clasp)
install:
    npm install

# Format every file in place
fmt:
    npx dprint fmt

# Type-check src/ against its JSDoc annotations
typecheck:
    npx tsc -p jsconfig.json

# Every check that must pass before pushing
check: fmt typecheck

# Log in to Google for clasp (one-off; token is stored in ~/.clasprc.json)
login:
    npx clasp login

# Show which local files clasp would push
status:
    npx clasp status

# Run all checks, then upload src/ to the Apps Script project
push: check
    npx clasp push
