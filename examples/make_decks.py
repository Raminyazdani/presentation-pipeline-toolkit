"""Build two tiny PowerPoint decks for the ordering example."""

import argparse
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches


def add_page(deck, title, description):
    slide = deck.slides.add_slide(deck.slide_layouts[5])
    slide.shapes.title.text = title
    body = slide.shapes.add_textbox(Inches(1), Inches(1.6), Inches(8), Inches(1))
    body.text_frame.text = description
    return slide


def mosaic_deck():
    deck = Presentation()
    deck.core_properties.title = "Tabletop mosaic"
    deck.core_properties.author = "Presentation Pipeline Toolkit"
    add_page(deck, "Pick three tiles", "Choose amber, blue and chalk tiles for a small sample.")
    slide = add_page(deck, "Count the colors", "One repeat uses ten tiles.")
    counts = CategoryChartData()
    counts.categories = ["Amber", "Blue", "Chalk"]
    counts.add_series("Tiles per repeat", (5, 3, 2))
    slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(2.6),
                           Inches(8), Inches(4), counts)
    add_page(deck, "Lay the border", "Reserve a narrow strip of chalk tiles for the edge.")
    return deck


def lantern_deck():
    deck = Presentation()
    deck.core_properties.title = "Paper lantern"
    deck.core_properties.author = "Presentation Pipeline Toolkit"
    add_page(deck, "Fold the shade", "Fold eight panels in a strip of pale green paper.")
    add_page(deck, "Join the ends", "Tuck the last flap into the first fold to close the ring.")
    return deck


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory", type=Path, help="new output directory")
    directory = parser.parse_args().directory
    try:
        directory.mkdir(parents=True)
        for name, build in (("mosaic", mosaic_deck), ("lantern", lantern_deck)):
            target = directory / f"{name}.pptx"
            with target.open("xb") as stream:
                build().save(stream)
            print(f"Created {target}")
    except OSError as exc:
        parser.exit(2, f"make_decks: {exc}\n")


if __name__ == "__main__":
    main()
