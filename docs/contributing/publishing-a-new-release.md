# Publishing a new release

Let's create a release from time to time, every 1 or 2 months for instance, to contain enough changes, but not too much.
Before beginning, you should ensure that ZDEMO_EXCEL_CHECKER (in the demos repo) shows all green checkmarks.

Below are the notes taken while publishing the release `7.16.0`.

Version numbering is based on [Semantic Versioning 2.0.0](https://semver.org/):
- `7`: a major release. NB: it will probably not change as we don't want to "make incompatible API changes".
- `16`: a minor release
- `0`: patch level (bug fixes)

Working directly on the upstream repository:
- create a branch for this new release; suggested naming for the branch: your own prefix - slash - release - number. For example: abo/release7.16.0 OR sandraros/release7.16.0
- change `version` in `zcl_excel` to indicate the new version number
- push the changes to this new release branch

With GitHub Desktop (or any Git console or Git user interface), [add the (lightweight) tag](https://docs.github.com/en/desktop/contributing-and-collaborating-using-github-desktop/managing-commits/managing-tags) `v7.16.0` to this branch; suggested naming for version-related tags is v + version number.

Do a pull request.

Wait for approval/commit(s) merged into the master branch.

Now [create the release in GitHub](https://docs.github.com/en/repositories/releasing-projects-on-github/managing-releases-in-a-repository#creating-a-release):
- Click "Releases"
- Click "Draft a new release"
- Click "Choose a tag"
- Type the title, select the previous tag, click "Auto-generate release notes" and click "Preview" to verify; you should have a list with the changes from the previous release, edit as required and remember to include the list below as explanation: 
    - `+`: new feature
    - `*`: bug fix
    - `!`: feature modification
    - `-`: feature removed
- Click "Publish release" 
- It's done, zip and tar.gz files are automatically assigned to the release
- The new release appears in the Code home page, with the changelog in the release page itself.

Now create a release for the [demos](https://github.com/abap2xlsx/demos) repository as well: use the same process and tag the latest commit available at the time the main library is released, to ensure users will always have a matching set of demo programs.

