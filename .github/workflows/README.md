# GitHub Actions Workflows

This directory contains automated CI/CD workflows for the Synth Riders DiscordRPC project.

## test-and-build.yml

Automated testing and building workflow that runs on all pushes and pull requests.

### When It Runs

- **Push to any branch** - Automatically tests and builds
- **Pull requests to main/master** - Tests and builds PR changes
- **Manual trigger** - Can be triggered manually from GitHub Actions tab

### What It Does

#### 1. Test Job
- Runs on Windows latest (free runner)
- Sets up Python 3.10
- Installs all dependencies (cached for speed)
- Runs the full pytest test suite (37 tests)
- Uploads test results as artifacts

#### 2. Build Job (only if tests pass)
- Runs on Windows latest (free runner)
- Installs dependencies
- Executes `build.bat` to create executables
- Generates BUILD_INFO.txt with version and commit details
- Uploads build artifacts with smart naming:
  - **Main/Master branch**: `synth-riders-discordrpc-v{version}-{sha}` (90-day retention)
  - **Other branches**: `synth-riders-discordrpc-{branch}-{sha}` (30-day retention)

### Artifacts

After a successful workflow run, you can download:

1. **test-results** - JUnit XML test results (30-day retention)
2. **Build artifacts** - All three executables plus BUILD_INFO.txt
   - `Synth Riders DiscordRPC.exe` - Main application
   - `Uninstall Synth Riders DiscordRPC.exe` - Uninstaller
   - `Synth Riders DiscordRPC Setup.exe` - Installer/setup

### How to Download Artifacts

1. Go to the [Actions tab](../../actions) in GitHub
2. Click on a workflow run
3. Scroll down to "Artifacts" section
4. Click on the artifact name to download

### Testing the Workflow

#### First Time Setup

1. Push the `optimize-maintainability` branch to GitHub:
   ```bash
   git push origin optimize-maintainability
   ```

2. Go to GitHub → Actions tab
3. You should see the workflow running

#### View Results

- **Green checkmark** ✅ - Tests passed, build succeeded
- **Red X** ❌ - Tests or build failed
- Click on the workflow run to see detailed logs

#### Download Testing Builds

1. Navigate to Actions → Select a workflow run
2. Scroll to "Artifacts" section
3. Download the build artifact (e.g., `synth-riders-discordrpc-optimize-maintainability-abc1234`)
4. Extract and test the executables

### Status Badge

Add this to your README.md to show build status:

```markdown
[![Test and Build](https://github.com/6uhrmittag/Synth-Riders-DiscordRPC/actions/workflows/test-and-build.yml/badge.svg)](https://github.com/6uhrmittag/Synth-Riders-DiscordRPC/actions/workflows/test-and-build.yml)
```

### Caching

The workflow uses caching to speed up runs:
- **Pip packages** - Cached by requirements.txt/requirements-test.txt hash
- Significantly reduces dependency installation time on subsequent runs

### Windows-Specific Features

- Uses `windows-latest` runner (includes Windows Server 2022)
- PowerShell scripts for Windows-specific tasks
- Proper path handling for Windows (`~\AppData\Local\pip\Cache`)
- Native Windows executable building with PyInstaller

### Troubleshooting

#### Workflow doesn't appear
- Make sure the workflow file is committed
- Check that you're looking at the correct branch
- Wait a moment - workflows may take a few seconds to appear

#### Tests fail
- Check the test results artifact for details
- View the workflow logs for full pytest output
- Tests should pass locally before pushing

#### Build fails
- Ensure `build.bat` works locally
- Check that all dependencies are in `requirements.txt`
- Review PyInstaller output in workflow logs

#### Artifacts not uploading
- Check that `dist/` folder contains executables
- Verify the artifact naming logic in the workflow
- Ensure paths are correct for Windows

### Cost

- ✅ **Free for public repositories**
- Uses GitHub-hosted Windows runners
- No additional configuration needed
- 2,000 minutes/month free for private repos (if needed)

### Future Enhancements

Potential additions to this workflow:

- [ ] Automatic release creation for tagged commits
- [ ] Code coverage reporting
- [ ] Linting checks (flake8, black, mypy)
- [ ] Security scanning
- [ ] Deploy to release page on main branch
- [ ] Notification on build failure

### Manual Workflow Trigger

You can manually trigger the workflow:

1. Go to Actions → test-and-build workflow
2. Click "Run workflow" button
3. Select branch and run
4. Useful for re-running builds without new commits

## Benefits

✅ **Automated Testing** - Catch bugs before merging  
✅ **Cross-Branch Builds** - Test executables from any branch  
✅ **Free CI/CD** - No cost for public repos  
✅ **Fast Feedback** - Know immediately if changes break anything  
✅ **Easy Testing** - Download and test builds without local setup  
✅ **Quality Assurance** - Tests must pass before builds  

---

For more information about GitHub Actions, see:
- [GitHub Actions Documentation](https://docs.github.com/en/actions)
- [Workflow syntax](https://docs.github.com/en/actions/using-workflows/workflow-syntax-for-github-actions)

