import mathml2latex, importlib.metadata, inspect, sys
from mathml2latex import convert as mathml2latex

print("Loaded from :", mathml2latex.__file__)
try:
    print("Version     :", importlib.metadata.version("mathml2latex"))
except importlib.metadata.PackageNotFoundError:
    print("Version     : (no dist-info, probably a local file)")
print("Has convert?:", hasattr(mathml2latex, "convert"))
