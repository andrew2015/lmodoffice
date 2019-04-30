# lmodoffice

A simple [Lua](https://github.com/ofunc/lua) module for converting various office documents into OOXML format files.

## Usage

```go
package main

import (
	"ofunc/lmodoffice"
	"ofunc/lua/util"
)

func main() {
	l := util.NewState()
	l.Preload("office", lmodoffice.Open)
	util.Run(l, "main.lua")
}
```

```lua
local office = require 'office'

office.toxlsx('path to files')
```

## Dependencies

* [ofunc/lua](https://github.com/ofunc/lua)
* [go-ole/go-ole](https://github.com/go-ole/go-ole)

## Documentation

### office.toxlsx(root[, ext1, ...])

Converts all files in the root that has the specified extensions to xlsx format files.
If no extension is specified, converts all files in the root.

### office.todocx(root[, ext1, ...])

Converts all files in the root that has the specified extensions to docx format files.
If no extension is specified, converts all files in the root.

### office.topptx(root[, ext1, ...])

Converts all files in the root that has the specified extensions to pptx format files.
If no extension is specified, converts all files in the root.
