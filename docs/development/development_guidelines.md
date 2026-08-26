# Development Guidelines
## Acceptable Submissions
Currently, only submissions from City of Philadelphia Employees will be considered.

To maintain code quality, security, and maintainability, this project prohibits "vibe-coded" contributions. Here, we mean "vibe-coded" as the submission of AI generated code that the author cannot validate or explain.

While contributions may be made with AI assistance, we expect all contributors to understand and be able to explain every line of code they submit. 

Contributions must be tested, and the results of those tests must be communicated before acceptance.

Contributions making substantive changes (updates to call signatures, added functions, changes to data returned by existing functions, major changes to function logic, etc.) should include documentation revisions.

## Testing
Small edits (e.g., consistency, code cleanup, etc) may not require addtional pytests. Larger edits may require pytests covering new functionality. The Repo maintainers may send your submission back for further edits if they find that test coverage is not sufficient.

Please run pytest before submission to ensure that your updates do not break existing tests.

For substantive edits, simply running pytest is not enough. We expect submissions to be tested and validated, and a description of those tests to be provided with any PR.

## Code style and architecture
Before submitting, read through the repo to get a sense of how the code is structured. Note the use of Pydantic models for API response validation.

Use type hints. This helps with debugging, and helps dev tools like `mypy` identify potential bugs in the code.

## Documentation

Docstrings should follow Google style. For more information on Google style, **[you can refer to Google's public guide.](https://google.github.io/styleguide/pyguide.html#38-comments-and-docstrings)**

Generally, doc-strings should take the following format:
```
"""
One line summary.

Extended multi-line description, if necessary.

Args:
    arg_one: Short description of argument.

Returns:
    Short description of returned data.

Raises:
    If function may raise an exception, describe the exception.
"""
```

Full documentation is created using Zensical, which serves markdown stored in the `/docs/` folder at the top level of this repo. This includes automated documentation of docstrings using mkdocstrings, which are stored in the `/reference` section of the documentation.. When writing new user-facing functions or classes, please review and update the relevant documentation, and include the automated docstrings in the reference section.