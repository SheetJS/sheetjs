#### Column Properties

The `!cols` array in each worksheet, if present, is a collection of `ColInfo`
objects which have the following properties:

```typescript
type ColInfo = {
	/* visibility */
	hidden:?boolean; // if true, the column is hidden

	/* column width is specified in one of the following ways: */
	wpx?:number;     // width in screen pixels
	width?:number;    // width in Excel's "Max Digit Width", width*256 is integral
	wch?:number;     // width in characters

	/* other fields for preserving features from files */
	MDW?:number;     // Excel's "Max Digit Width" unit, always integral
};
```

Excel internally stores column widths in a nebulous "Max Digit Width" form.  The
Max Digit Width is the width of the largest digit when rendered (generally the
"0" character is the widest).  The internal width must be an integer multiple of
the the width divided by 256.  ECMA-376 describes a formula for converting
between pixels and the internal width.

<details>
	<summary><b>Implementation details</b> (click to show)</summary>

Given the constraints, it is possible to determine the MDW without actually
inspecting the font!  The parsers guess the pixel width by converting from width
to pixels and back, repeating for all possible MDW and selecting the MDW that
minimizes the error.  XLML actually stores the pixel width, so the guess works
in the opposite direction.

Even though all of the information is made available, writers are expected to
follow the priority order:

1) use `width` field if available
2) use `wpx` pixel width if available
3) use `wch` character count if available
</details>

#### Row Properties

The `!rows` array in each worksheet, if present, is a collection of `RowInfo`
objects which have the following properties:

```typescript
type RowInfo = {
	/* visibility */
	hidden?:boolean; // if true, the row is hidden

	/* row height is specified in one of the following ways: */
	hpx?:number;     // height in screen pixels
	hpt?:number;     // height in points
};
```

<details>
	<summary><b>Implementation details</b> (click to show)</summary>

Excel internally stores row heights in points.  The default resolution is 72 DPI
or 96 PPI, so the pixel and point size should agree.  For different resolutions
they may not agree, so the library separates the concepts.

Even though all of the information is made available, writers are expected to
follow the priority order:

1) use `hpx` pixel height if available
2) use `hpt` point height if available
</details>

