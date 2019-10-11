# xlsxport

A simple python microservice to export data as an Excel xlsx spreadsheet.

While the excellent [`js-xlsx`](https://github.com/SheetJS/js-xlsx) project has
a rich feature set and allows client-side exporting, the community edition
doesn't include style configuration. Also, offloading some work to a
service is kind to your users' devices and batteries.

## Usage

Accepts an array of arrays (rows and columns) of data via POST body, and returns an Excel
xlsx spreadsheet with the data.

### Primitive values

If an input cell's value is a primitive, the output cell will simply contain
the value. For example, the following input will behave how you expect:

```json
[
	[1, 2, 3],
	["a", "b", "c"]
]
```

### Objects

If an object, the following properties are used:

- `value`: The output cell's value
- `style`: The cell's style. Currently only builtin string styles are
supported.

For example, the following input will result in the cell `A2` containing `2`
and being highlighted with Excel's builtin `Accent1` style:

```json
[
	[1, { "value": 2, "style": "Accent1" }, 3]
]
```

More configuration options, such as formatting, more advanced styles, and
formulas are planned and will be added as needed.

## Deployment and configuration

Everything is handled in `index.py`, which exports a `handler` class that
extends `BaseHTTPRequestHandler`. Designed for simple deployment with
[Zeit's `now`](https://zeit.co/), simply modifying the `now.json` configuration
file to your desires and running `now` should work out of the box.

CORS headers and preflight requests should be handled appropriately out of the
box using the ALLOWED_ORIGINS environment variable, a space-separated list of
domains that should be allowed to use the service. Omitting this, or setting it
to `*` allows any domain to use the service.

