# XBRL Taxonomy Package Conformant Processor

## :newspaper: About the project

The script allows the user to automatically generate license approval forms for a company.

### Content overview

    .
    ├── img/ - contains company logo template
    ├── templates/ - folder with data templates
    ├── YYYY-MM-DD/ - output folder for generated licenese approval forms
    ├── .gitignore - list of files/fodlers not tracked by git
    ├── Constants.py - contain relevant data for the license approval generation
    ├── gen_lic_approval.py - drving code for the license approval form generation
    ├── LICENSE - license text of project
    └── README.md - contains project information

## :notebook: Features

* Generate license approval in DOCX format.

## :runner: Getting started

### Example usage

```python
gen_lic_approval.py [family] [version]
```

```python
gen_lic_approval.py [-family='EBA'] [-version="3.2"]
```

## :books: Resources used to create this project

* Python
  * [Python 3.10.13 documentation](https://docs.python.org/3.10/)
  * [Built-in Functions](https://docs.python.org/3.10/library/functions.html)
  * [Python Module Index](https://docs.python.org/3.10/py-modindex.html)
* Markdwon
  * [Basic syntax](https://www.markdownguide.org/basic-syntax/)
  * [Complete list of github markdown emofis](https://dev.to/nikolab/complete-list-of-github-markdown-emoji-markup-5aia)
  * [Awesome template](http://github.com/Human-Activity-Recognition/blob/main/README.md)
  * [.gitignore file](https://git-scm.com/docs/gitignore)
* Editor
  * [Visual Studio Code](https://code.visualstudio.com/)

## :bookmark: License

[GPL v3](https://www.gnu.org/licenses/gpl-3.0.txt) :copyright: 2023 Christoph Hartleb
