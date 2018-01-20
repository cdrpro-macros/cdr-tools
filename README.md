# CdrTools

A source code of CdrTools macro.

## About CdrTools

CdrTools is a large collection of simple macros for daily work in CorelDRAW.
The package includes about 40 functions for working with the clipboard,
colors, documents and objects.

* Version: 2.11 (2013)
* Compatibility: X3 – X6

### Functions:

#### sToolsAbout

* ShowAbout - About dialog.
* OpenOptionsWindow - Options.

#### sToolsBitmap

* ToPowerClip - Place each bitmap from active selection to PowerClip specially created for it.

#### sToolsClipboard

* Clear - clear the clipboard.
* PasteInCenter - paste in center of active view.
* PasteAsBitmap / PasteAsMetafile / PasteAsCorel32 / PasteAsText - fast paste as...
* PasteOrderBack - paste order back from selection.
* PasteOrderFront - paste order front from selection.

#### sToolsColor

* ToRGB / ToCMYK / ToGRAY - fast convert all colors (vector & bitmap) to ...
* ReConvertCMYK - reconvert all colors cmyk to cmyk.
* EditStepsNumOfFill - change steps in fountain fill and linear transparency.
* SelectSame - select objects with the same uniform fill (hold Ctrl for select with
  the same outline color).

#### sToolsDocument

* SwitchView - switch view (between enhanced/pixels & wireframe).
* NewWebDocument - create new web 72 dpi document (one click).
* NewCMYKDocument - create new cmyk 300 dpi document (one click).
* GoToFirstPage / GoToLastPage - go to first/last page.
* FitPageToSelect - fit page size to size of selection.
* DeleteViewStyles - delete all view styles.
* RemovePagesNames - delete names of all pages (shows numbers only).

#### sToolsOutline

* MakeOutlineSameAsFill - create outline with the same color as fill
  (useful to handle the result of the trace).
* SetScale / SetNoScale - on/off scale outline.
* OutlineWidthUp / OutlineWidthDown - increase/decrease outline width.

#### sToolsShape

* LineSpacingUp / LineSpacingDown - increase/decrease line spacing for text
  (step value can be changed in settings).
* RotateRight / RotateLeft - rotate selection (step value can be changed in settings).
  Hold Shift for using second value.
* MoveToDesktop - move selection objects to Desktop layer.
* MoveToCenter - move selection objects to center of view.
* SmartIntersect - intersect (delete front object).
* DeleteNoCloseCurves - delete all open curves in selection.
* DeArtBrush - convert all brushes to objects and delete path in selection.
* SwapTwoShapes - swap position of two objects (hold Shift - swap position and size)
* GetCurveInfo - show information (area, length, number of nodes...) about selection curve.
* Oculist - resize letters by ascending, descending, and randomly.

#### sToolsTest

* SpeedTest - simple test of speed your CorelDRAW & PC
  (result can be compared with other CorelDRAW/PC).

#### sToolsWorkspace

* SwitchRulersGuidelines - one click for change state of 'Edit > Rulers' and 'Edit > Guidelines'.

#### Additional features

* Auto Update Text (for X6 only) - If you don't want to see 'Update Text' panel
  when you open cdr files, just enable this macro in options windows of CdrTools.
  Also, don't forget disable 'Delay Load VBA' in Tools > Options > Workspace > VBA.

## License

Copyright © 2018 [Sancho](http://cdrpro.ru/en/)

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see http://www.gnu.org/licenses/.