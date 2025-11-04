# Publishing Lokka to NPM

## Current Version

**0.3.1** - Last published version

## Publishing Checklist

### 1. Pre-Publishing Steps

- [ ] **Test locally**: Ensure all tests pass

  ```powershell
  cd c:\temp\work\lokka\src\mcp
  npm run build
  npm run test:token  # Or other tests
  ```

- [ ] **Update version**: Decide on version bump type

  - **Patch** (0.3.1 → 0.3.2): Bug fixes, small changes
  - **Minor** (0.3.1 → 0.4.0): New features, backward compatible (REST API wrapper = minor!)
  - **Major** (0.3.1 → 1.0.0): Breaking changes

- [ ] **Update CHANGELOG**: Document what's new (if you have one)

- [ ] **Commit all changes**: Ensure working directory is clean
  ```powershell
  git status
  git add .
  git commit -m "Release v0.4.0"
  ```

### 2. Version Bump

Choose one method:

#### Method A: Using npm version (Recommended)

```powershell
cd c:\temp\work\lokka\src\mcp

# For REST API wrapper addition (new feature = minor version)
npm version minor  # 0.3.1 → 0.4.0

# Or for bug fixes only
npm version patch  # 0.3.1 → 0.3.2

# Or for breaking changes
npm version major  # 0.3.1 → 1.0.0
```

This automatically:

- Updates package.json
- Creates a git commit
- Creates a git tag

#### Method B: Manual version update

Edit `package.json` and change version manually, then:

```powershell
git add package.json
git commit -m "Bump version to 0.4.0"
git tag v0.4.0
```

### 3. Build the Package

```powershell
cd c:\temp\work\lokka\src\mcp

# Clean and rebuild
Remove-Item -Recurse -Force build -ErrorAction SilentlyContinue
npm run build
```

### 4. Test the Package Locally (Optional)

Test that the package works when installed:

```powershell
# Create a test directory
cd c:\temp
mkdir test-lokka
cd test-lokka

# Install from local path
npm init -y
npm install ..\work\lokka\src\mcp

# Test it
npx lokka --help
```

### 5. Login to NPM

```powershell
# Login to npm (if not already logged in)
npm login

# Verify you're logged in
npm whoami
```

### 6. Publish to NPM

```powershell
cd c:\temp\work\lokka\src\mcp

# Dry run first (see what will be published)
npm publish --dry-run

# Actually publish
npm publish --access public
```

**Note**: Since this is a scoped package (`@thiagobeier/lokka`), you need `--access public` unless you have a paid npm account.

### 7. Push to GitHub

```powershell
# Push commits
git push origin main

# Push tags
git push origin --tags
```

### 8. Verify Publication

```powershell
# Check on npm
npm view @thiagobeier/lokka

# Try installing from npm
npm install -g @thiagobeier/lokka

# Test the installed version
lokka --help
```

### 9. Create GitHub Release (Optional)

1. Go to https://github.com/thiagogbeier/lokka/releases
2. Click "Draft a new release"
3. Choose the tag you created (e.g., v0.4.0)
4. Add release notes describing the new REST API wrapper feature
5. Publish release

## Quick Publish Script

For future releases, you can use this quick script:

```powershell
# Quick publish (after ensuring all changes are committed)
cd c:\temp\work\lokka\src\mcp

# Clean build
Remove-Item -Recurse -Force build -ErrorAction SilentlyContinue
npm run build

# Version bump (choose one)
npm version minor  # For new features like REST API wrapper

# Publish
npm publish --access public

# Push to GitHub
git push origin main --tags
```

## Common Issues

### "You must be logged in to publish packages"

```powershell
npm login
```

### "You do not have permission to publish"

Ensure you're logged in as the package owner:

```powershell
npm whoami  # Should show: thiagobeier
```

### "Package already exists"

The version you're trying to publish already exists. Bump the version:

```powershell
npm version patch
```

### "This package has been marked as private"

Remove `"private": true` from package.json or add `--access public`

## What Gets Published

Based on your `package.json`, these files will be included:

- `build/` directory (containing compiled JavaScript)
- `package.json`
- `README.md` (from src/mcp directory)
- `LICENSE` (if present)

The `.dockerignore` file won't be included, but npm has its own ignore rules.

## Version History

- **0.3.1**: Current published version
- **0.4.0**: (Planned) Add REST API wrapper for Copilot Studio integration

## Recommended Version for This Update: 0.4.0

Since you're adding a **major new feature** (REST API wrapper), this warrants a **minor version bump** (0.3.1 → 0.4.0).

Changes in this release:

- ✨ New REST API wrapper for Copilot Studio integration
- 🐳 Docker support for Azure Container Apps
- 📝 OpenAPI specification for API documentation
- 🔧 Deployment automation scripts
- 📚 Comprehensive deployment documentation

## Publishing Command Summary

```powershell
cd c:\temp\work\lokka\src\mcp

# 1. Build
npm run build

# 2. Version bump
npm version minor  # Creates 0.4.0

# 3. Publish
npm publish --access public

# 4. Push to GitHub
git push origin main --tags
```
