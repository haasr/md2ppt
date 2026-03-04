Markdown-2-PowerPoint (md2ppt)
==============================

Converts a Markdown file into a ``.pptx`` PowerPoint presentation using ``python-pptx``.

Sometimes, you need to make a PowerPoint quickly and you already have the material -- it just isn't in slides format.
This package makes it easy to convert a Markdown file into basic PowerPoint slides and sticks with conventional
PowerPoint layouts so you can use the "Designer" feature to instantly style the slides.

**Why Markdown?** Because every LLM can write Markdown, so as long as it follows the conventions below,
it can be piped directly into ``SlidesBuilder`` to produce a slide deck.

Installation
------------

.. code-block:: bash

    pip install md2ppt

CLI Usage
---------

.. code-block:: bash

    md2ppt input.md output.pptx

The CLI applies a built-in ETSU color theme by default.

Module Usage
------------

.. code-block:: python

    from md2ppt import SlidesBuilder

    with open("input.md") as f:
        md = f.read()

    SlidesBuilder(md, "output.pptx").build()

To apply a custom color theme, pass a ``theme_colors`` dict with any subset of Office theme slot names
(hex strings, no ``#``):

.. code-block:: python

    my_colors = {
        "accent1": "1A3C6E",   # primary brand blue
        "accent2": "E87722",   # accent orange
    }

    SlidesBuilder(md, "output.pptx", theme_colors=my_colors).build()

Valid slot names: ``dk1``, ``lt1``, ``dk2``, ``lt2``, ``accent1``–``accent6``, ``hlink``, ``folHlink``.

Markdown Conventions
--------------------

.. list-table::
   :header-rows: 1
   :widths: 30 70

   * - Syntax
     - Result
   * - ``# Title``
     - Title slide
   * - ``# Title: Subtitle``
     - Title slide with subtitle
   * - ``## Heading``
     - Content slide
   * - ``- item``
     - Bulleted list item
   * - ``1. item``
     - Numbered item (hanging indent)
   * - Plain paragraph text
     - Unbulleted body text

Example Input
-------------

.. code-block:: markdown

    # My Presentation: A Subtitle

    ## First Slide

    - Bullet point one
    - Bullet point two

    ## Second Slide

    1. First numbered item
    2. Second numbered item

    ## Thank You

    Contact us at example@example.com

Generating Slides-Compatible Markdown with an LLM
--------------------------------------------------

Use a prompt like the following to get an LLM to output Markdown that ``md2ppt`` can convert directly:

    Generate a slide deck in Markdown format on the topic of **"[Your Topic]"**.

    Follow these formatting rules exactly:

    - Use a single ``#`` heading for the title slide. If you want a subtitle, put it after a colon: ``# Title: Subtitle``
    - Use ``##`` headings for each content slide title
    - Use `-` bullet points for unordered lists
    - Use ``1.`` numbered lists for ordered/ranked content
    - Use plain paragraph text (no prefix) for prose body content
    - Do **not** use bold, italics, links, code blocks, or any other Markdown formatting
    - Aim for 5–8 slides total, with 4–6 items per content slide
    - Output **only** the Markdown — no explanation, no preamble

License
-------

MIT
