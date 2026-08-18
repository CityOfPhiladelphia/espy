# graph_api_functions

## Installation
### To install packages needed to run:
```
uv sync
```
### To install packages needed to develop this package:
```
uv sync --dev
```

## Adding dependencies
Before you add a dependency, ask: do I need this to **only to develop** the tool,
or is this needed to run the tool?

If you only need the package for development (for example, linters like ruff, pytest, ipython, etc...), then please add the --dev flag when adding the dependency:

```
uv add --dev <<package name>>
```

If the dependency is truly needed for functionality, run:
```
uv add <<package name>>
```

### When adding a dependency:
1. Consider whether or not the dependency is necessary before adding it at all. Python's library is quite large, and may have tools that perform the functionality you're looking for.
2. Research the dependency. Is it well maintained? How many stars does the GitHub repo have? When was the last time it was updated? Is there a better, more supported dependency out there that is less likely to be abandoned in the future?
3. Add dependencies through `uv`, not through `pip`.


## Importing and authenticating with the GraphAPIClient

The GraphAPIClient contains a class method that handles authentication
to the sharepoint API.


```
from graph_api_functions.client import GraphAPIClient

client = GraphAPIClient.authenticate()
```