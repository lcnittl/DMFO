"""
import pkgutil

__path__ = pkgutil.extend_path(__path__, __name__)

for ModuleInfo in pkgutil.walk_packages(path=__path__, prefix=__name__+'.'):
  _, module, _ = ModuleInfo
  __import__(module)

del pkgutil
del module
del ModuleInfo
del _
"""
"""
from pathlib import Path

__all__ = []

for path in Path(__file__).parent.iterdir():
    if path.is_file():
        path = Path(path.name)
        if path.name == '__init__.py' or path.suffix != '.py':
            continue
        module = path.stem
        __all__ += [module]
del Path
del path
del module
"""
