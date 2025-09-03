# Name card generator

A simple script to genreate printable front-and-back name cards.

## Install

Install [`reportlab`](https://pypi.org/project/reportlab/) with:

```bash
pip install reportlab
```
or
```bash
pip3 install reportlab
```
> Using a venv or conda environment is recommended.

## Usage

Edit the `name.txt` file in the format:
    name \n name ...

run 

```bash
python3 name_card.py
```

and a `name_badges.pdf` file would be generated. You can print it directly.

## Limitations

- No Error handling is involved. 
- Only support **Chinese name** with **2-4** characters.