# version from src/Consts.js
version := `grep -oE 'HUB_VERSION = "[^"]+"' src/Consts.js | cut -d'"' -f2`

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

# Delete dist/
clean:
    rm -rf dist

# Join all src files into dist/SpiritHub-<version>.js
build:
    node scripts/build.mjs

# Check, build, then push dist/SpiritHub-<version>.js to Apps Script
push: check build
    npx clasp push
    