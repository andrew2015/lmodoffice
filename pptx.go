/*
Copyright 2019 by ofunc

This software is provided 'as-is', without any express or implied warranty. In
no event will the authors be held liable for any damages arising from the use of
this software.

Permission is granted to anyone to use this software for any purpose, including
commercial applications, and to alter it and redistribute it freely, subject to
the following restrictions:

1. The origin of this software must not be misrepresented; you must not claim
that you wrote the original software. If you use this software in a product, an
acknowledgment in the product documentation would be appreciated but is not
required.

2. Altered source versions must be plainly marked as such, and must not be
misrepresented as being the original software.

3. This notice may not be removed or altered from any source distribution.
*/

package lmodoffice

import (
	"os"
	"path/filepath"
	"strings"

	"ofunc/lua"

	ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func lToPPTX(l *lua.State) int {
	ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED)
	defer ole.CoUninitialize()

	unknown, err := oleutil.CreateObject("PowerPoint.Application")
	if err != nil {
		l.Push(err.Error())
		return 1
	}
	powerpoint, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		l.Push(err.Error())
		return 1
	}
	defer powerpoint.Release()
	pres, err := powerpoint.GetProperty("Presentations")
	if err != nil {
		l.Push(err.Error())
		return 1
	}

	n := l.AbsIndex(-1)
	exts := make([]string, 0, n-1)
	for i := 2; i <= n; i++ {
		exts = append(exts, l.ToString(i))
	}
	app := pres.ToIDispatch()
	err = filepath.Walk(l.ToString(1), func(src string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if info.IsDir() {
			return nil
		}
		return toPPTX(app, src, exts)
	})
	if err != nil {
		l.Push(err.Error())
		return 1
	}
	_, err = powerpoint.CallMethod("Quit")
	if err != nil {
		l.Push(err.Error())
		return 1
	}
	return 0
}

func toPPTX(app *ole.IDispatch, src string, exts []string) error {
	ext := filepath.Ext(src)
	if ignore(ext, ".pptx", exts) {
		return nil
	}
	src, err := filepath.Abs(src)
	if err != nil {
		return err
	}
	tar := strings.TrimSuffix(src, ext) + ".pptx"

	pre, err := app.CallMethod("Open", src)
	if err != nil {
		return err
	}
	presentation := pre.ToIDispatch()
	_, err = presentation.CallMethod("SaveAs", tar, 24)
	if err == nil {
		_, err = presentation.CallMethod("Close")
	} else {
		presentation.CallMethod("Close")
	}
	if err != nil {
		return err
	}
	return os.Remove(src)
}
