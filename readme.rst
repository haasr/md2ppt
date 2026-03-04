Markdown-2-PowerPoint
=====================

Converts a Markdown file into a ``.pptx`` PowerPoint presentation using ``python-pptx``.

Sometimes, you need to make a PowerPoint quickly and you already have the material -- it just isn't in slides format.
This code makes it easy to convert a markdown file into basic PowerPoint slides and sticks with conventional
PowerPoint layouts so you can use the "Designer" feature to instantly style the slides.

**Why Markdown?** Because every LLM, lobotomized or not, can write markdown, so as long as that markdown follows
the `Markdown Conventions`_ below, it can be piped directly into the SlidesBuilder class to make a slideset.
See `The Prompt`_.

Requirements
------------

.. code-block:: bash

    pip install python-pptx lxml

CLI Usage
---------

.. code-block:: bash

    python md2ppt.py input.md output.pptx

The CLI applies a built-in ETSU color theme by default.

Module Usage
------------

.. code-block:: python

    from md2ppt import SlidesBuilder

    with open("input.md") as f:
        md = f.read()

    SlidesBuilder(md, "output.pptx").build()

To apply a custom color theme, pass a ``theme_colors`` dict with any subset of Office theme slot names (hex strings, no ``#``):

.. code-block:: python

    my_colors = {
        "accent1": "1A3C6E",   # primary brand blue
        "accent2": "E87722",   # accent orange
    }

    SlidesBuilder(md, "output.pptx", theme_colors=my_colors).build()

Valid slot names: ``dk1``, ``lt1``, ``dk2``, ``lt2``, ``accent1``–``accent6``, ``hlink``, ``folHlink``.

.. _Markdown Conventions:

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

Example Prompt for Generating Slides-Compatible Markdown
---------------------------------------------------------

Use a prompt like the one below to get an LLM to output Markdown that ``md2ppt.py`` can convert directly to a ``.pptx`` file.

.. _The Prompt:

The Prompt
~~~~~~~~~~

    Generate a slide deck in Markdown format on the topic of **"Why Every Developer Should Learn SQL"**.

    Follow these formatting rules exactly:

    - Use a single ``#`` heading for the title slide. If you want a subtitle, put it after a colon: ``# Title: Subtitle``
    - Use ``##`` headings for each content slide title
    - Use `-` bullet points for unordered lists
    - Use ``1.`` numbered lists for ordered/ranked content
    - Use plain paragraph text (no prefix) for prose body content like contact info or closing remarks
    - Do **not** use bold, italics, links, code blocks, horizontal rules, or any other Markdown formatting — these will not render in the slides
    - Aim for 5–8 slides total, with 4–6 items per content slide
    - Output **only** the Markdown — no explanation, no preamble

Example Output (What the LLM Should Return)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

.. code-block:: markdown

    # Why Every Developer Should Learn SQL: A Practical Case

    ## SQL Is Everywhere

    - Used in web apps, data pipelines, analytics, and embedded systems
    - Supported by PostgreSQL, MySQL, SQLite, SQL Server, and more
    - The language has been largely unchanged since 1986 — your skills don't expire
    - Most backend frameworks still rely on relational databases under the hood

    ## You Will Write It Whether You Plan To or Not

    - ORMs abstract SQL but don't eliminate it — you still debug raw queries
    - Performance tuning requires reading query execution plans
    - Data migrations, reporting, and admin tasks almost always require raw SQL
    - Knowing SQL makes you less dependent on your data team

    ## Top 5 SQL Concepts Worth Mastering

    1. JOINs - Combining data across related tables
    2. Aggregations - GROUP BY, COUNT, SUM, AVG
    3. Window Functions - Running totals, rankings, and partitioned analytics
    4. Indexes - Understanding why your queries are slow
    5. Transactions - Keeping data consistent when things go wrong

    ## Getting Started

    Free resources to begin learning today:

    - SQLZoo and Mode Analytics offer free interactive tutorials
    - Practice on any SQLite database — no server required
    - Try querying a dataset you actually care about

    ## Conclusion

    SQL is one of the highest-return skills a developer can add.
    Start with SELECT. The rest follows naturally.