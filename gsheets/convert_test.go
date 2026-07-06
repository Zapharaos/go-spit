package gsheets

import (
	"testing"

	spit "github.com/Zapharaos/go-spit"
)

func TestColumnLetter(t *testing.T) {
	cases := map[int]string{1: "A", 2: "B", 26: "Z", 27: "AA", 52: "AZ", 53: "BA"}
	for in, want := range cases {
		if got := columnLetter(in); got != want {
			t.Errorf("columnLetter(%d) = %q, want %q", in, got, want)
		}
	}
}

func TestHexColor(t *testing.T) {
	c := hexColor("#FF0000")
	if c == nil || c.Red != 1 || c.Green != 0 || c.Blue != 0 || c.Alpha != 1 {
		t.Errorf("hexColor(#FF0000) = %+v", c)
	}
	if hexColor("nope") != nil {
		t.Error("invalid hex should return nil")
	}
	// Bare (no #) 6-digit hex is accepted.
	if got := hexColor("00FF00"); got == nil || got.Green != 1 {
		t.Errorf("hexColor(00FF00) = %+v", got)
	}
}

func TestBorderStyle(t *testing.T) {
	cases := map[spit.BorderStyle]string{
		spit.BorderStyleThin:   "SOLID",
		spit.BorderStyleMedium: "SOLID_MEDIUM",
		spit.BorderStyleThick:  "SOLID_THICK",
		spit.BorderStyleDashed: "DASHED",
		spit.BorderStyleDotted: "DOTTED",
		spit.BorderStyleDouble: "DOUBLE",
	}
	for in, want := range cases {
		if got := borderStyle(in); got != want {
			t.Errorf("borderStyle(%v) = %q, want %q", in, got, want)
		}
	}
}

func TestExtendedValue(t *testing.T) {
	if v := extendedValue(42); v == nil || v.NumberValue == nil || *v.NumberValue != 42 {
		t.Errorf("int -> %+v, want number 42", v)
	}
	if v := extendedValue(3.5); v == nil || v.NumberValue == nil || *v.NumberValue != 3.5 {
		t.Errorf("float -> %+v, want number 3.5", v)
	}
	if v := extendedValue("hi"); v == nil || v.StringValue == nil || *v.StringValue != "hi" {
		t.Errorf("string -> %+v, want string hi", v)
	}
	if v := extendedValue(true); v == nil || v.BoolValue == nil || *v.BoolValue != true {
		t.Errorf("bool -> %+v, want bool true", v)
	}
	if v := extendedValue(""); v != nil {
		t.Errorf("empty string -> %+v, want nil", v)
	}
	if v := extendedValue(nil); v != nil {
		t.Errorf("nil -> %+v, want nil", v)
	}
}
