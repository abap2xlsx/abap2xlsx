# Publishing a new release

Let's create a release from time to time, every 1 or 2 months for instance, to contain enough changes, but not too much.

Below are the notes taken while publishing the release `7.15.0`.

Version numbering is based on [Semantic Versioning 2.0.0](https://semver.org/):
- `7`: a major release. NB: it will probably not change as we don't want to "make incompatible API changes".
- `15`: a minor release
- `0`: patch level (bug fixes)

Create a branch for this new release, change `version` in `zcl_excel` to indicate the new version number and push the changes to the repository

With GitHub Desktop (or any Git console or Git user interface), [add the tag](https://docs.github.com/en/desktop/contributing-and-collaborating-using-github-desktop/managing-commits/managing-tags) `7.15.0` to this branch.

Do a pull request.

Wait for approval/commit(s) merged into the master branch.

Now [create the release in GitHub](https://docs.github.com/en/repositories/releasing-projects-on-github/managing-releases-in-a-repository#creating-a-release):
- Click "Releases"
- Click "Draft a new release"
- Click "Choose a tag"
- Type the title, click "Auto-generate release notes" and click "Preview" to verify; you should have a list with the following items, edit as required:
    - `+`: new feature
    - `*`: bug fix
    - `!`: feature modification
    - `-`: feature removed
- Click "Publish release" 
- It's done, zip and tar.gz files are automatically assigned to the release
- The new release appears in the Code home page, with the changelog in the release page itself.
