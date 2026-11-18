# Presentation Pipeline Toolkit

`pptx-order` changes the normal reading order of a PowerPoint deck. Give it a complete list of slide positions, starting at 1, and a new output path. It moves the slide references in the presentation manifest and copies every other package part intact. Charts stay editable and speaker notes stay attached to their slides.

## Quick start

Use Python 3.10 or newer. In a checkout of this repository, create a virtual environment and install the CLI with its example dependencies.

```sh
python -m venv .venv
. .venv/bin/activate
python -m pip install '.[demo]'
python examples/make_decks.py demo
pptx-order demo/mosaic.pptx --order 3 1 2 -o demo/ordered.pptx
```

On Windows PowerShell, activate the environment with `.venv\Scripts\Activate.ps1`. The generator creates a fresh directory containing a three-slide mosaic guide with an editable column chart and a two-slide paper lantern guide. Use a new directory when you repeat the demo.

The CLI prints `Wrote demo/ordered.pptx (3 slides in order 3 1 2)`. Reopen the output to read its slide titles.

```sh
python -c "from pptx import Presentation; print([s.shapes.title.text for s in Presentation('demo/ordered.pptx').slides])"
```

```text
['Lay the border', 'Pick three tiles', 'Count the colors']
```

To reverse the lantern guide, run `pptx-order demo/lantern.pptx --order 2 1 -o demo/lantern-reversed.pptx`. For CLI-only use, install with `python -m pip install .`.

## Input and output

The input must be an ordinary, unencrypted Transitional `.pptx` with its presentation part at `ppt/presentation.xml`. Use an existing output directory on a filesystem that supports hard links. The CLI publishes the finished file with an exclusive link. An existing output, including the input itself, produces an error. Missing, repeated and out-of-range positions also produce errors and exit with status 2.

Digitally signed packages and decks with PowerPoint sections are rejected. Custom shows retain their own sequences. Hand-authored agenda text, embedded slide numbers, metadata and cached thumbnails retain their original values. The operation preserves embedded content, including speaker notes and media, when you share the resulting file.

## Check the build and behavior

With the demo dependencies installed, run the test suite and build the wheel.

```sh
python -m unittest discover -s tests -v
python -m pip wheel --no-deps . -w dist
```

The tests reopen both reordered guides, check the chart's data and slide placement, compare every package part outside the manifest, and exercise validation and overwrite protection.

## License

The CLI, generator, tests and authored examples use the MIT license. Installed dependencies retain the licenses and notices shipped with their distributions. The CLI uses lxml under the BSD license. The demo uses python-pptx under the MIT license. [Project license](LICENSE) · [lxml license](https://lxml.de/copyright.html) · [python-pptx license](https://github.com/scanny/python-pptx/blob/master/LICENSE)
