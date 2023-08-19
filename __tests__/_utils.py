import typing

import click.testing

RUNNER_TYPE = typing.Callable[[typing.List[str]], click.testing.Result]
