var XLSX = require('xlsx');
var OUTFILE = '/tmp/example-style.xlsx';

function JSDateToExcelDate(inDate) {
  return 25569.0 + ((inDate.getTime() - (inDate.getTimezoneOffset() * 60 * 1000)) / (1000 * 60 * 60 * 24));
}

var defaultCellStyle = { font: { name: "Verdana", sz: 11, color: "FF00FF88"}, fill: {fgColor: {rgb: "FFFFAA00"}}};

// test to see if everything on the left equals its counterpart on the right
// but the right hand object may have other attributes which we might not care about
function basicallyEquals(left, right) {
  if (Array.isArray(left) && Array.isArray(right)) {
    for (var i = 0; i < left.length; i++) {
      if (!basicallyEquals(left[i], right[i])) {
        return false;
      }
    }
    return true;
  }
  else if (typeof left == 'object' && typeof right == 'object') {
    for (var key in left) {
      if (key != 'bgColor') {
        if (!basicallyEquals(left[key], right[key])) {
          if (JSON.stringify(left[key]) == "{}" && right[key] == undefined) return true;
          if (JSON.stringify(right[key]) == "{}" && left[key] == undefined) return true;
          return false;
        }
      }
    }
    return true;
  }
  else {
    if (left != right) {
      return false;
    }
    return true;
  }
}


var workbook, wbout, wbin;

workbook = {
  "SheetNames": [
    "Main"
  ],
  "Sheets": {
    "Main": {
      "!merges": [
        {
          "s": {
            "c": 0,
            "r": 0
          },
          "e": {
            "c": 2,
            "r": 0
          }
        }
      ],
      "A1": {
        "v": "This is a submerged cell",
        "s": {
          "border": {
            "left": {
              "style": "thick",
              "color": {
                "auto": 1
              }
            },
            "top": {
              "style": "thick",
              "color": {
                "auto": 1
              }
            },
            "bottom": {
              "style": "thick",
              "color": {
                "auto": 1
              }
            }
          }
        },
        "t": "s"
      },
      "B1": {
        "v": "Pirate ship",
        "s": {
          "border": {
            "top": {
              "style": "thick",
              "color": {
                "auto": 1
              }
            },
            "bottom": {
              "style": "thick",
              "color": {
                "auto": 1
              }
            }
          }
        },
        "t": "s"
      },
      "C1": {
        "v": "Sunken treasure",
        "s": {
          "border": {
            "right": {
              "style": "thick",
              "color": {
                "auto": 1
              }
            },
            "top": {
              "style": "thick",
              "color": {
                "auto": 1
              }
            },
            "bottom": {
              "style": "thick",
              "color": {
                "auto": 1
              }
            }
          }
        },
        "t": "s"
      },
      "A2": {
        "v": "Blank",
        "t": "s"
      },
      "B2": {
        "v": "Red",
        "s": {
          "fill": {
            "fgColor": {
              "rgb": "FFFF0000"
            }
          }
        },
        "t": "s"
      },
      "C2": {
        "v": "Green",
        "s": {
          "fill": {
            "fgColor": {
              "rgb": "FF00FF00"
            }
          }
        },
        "t": "s"
      },
      "D2": {
        "v": "Blue",
        "s": {
          "fill": {
            "fgColor": {
              "rgb": "FF0000FF"
            }
          }
        },
        "t": "s"
      },
      "E2": {
        "v": "Theme 5",
        "s": {
          "fill": {
            "fgColor": {
              "theme": 5
            }
          }
        },
        "t": "s"
      },
      "F2": {
        "v": "Theme 5 Tint -0.5",
        "s": {
          "fill": {
            "fgColor": {
              "theme": 5,
              "tint": -0.5
            }
          }
        },
        "t": "s"
      },
      "A3": {
        "v": "Default",
        "t": "s"
      },
      "B3": {
        "v": "Arial",
        "s": {
          "font": {
            "name": "Arial",
            "sz": 24,
            "color": {
              "theme": "5"
            }
          }
        },
        "t": "s"
      },
      "C3": {
        "v": "Times New Roman",
        "s": {
          "font": {
            "name": "Times New Roman",
            bold: true,
            underline: true,
            italic: true,
            strike: true,
            outline: true,
            shadow: true,
            vertAlign: "superscript",
            "sz": 16,
            "color": {
              "rgb": "FF2222FF"
            }
          }
        },
        "t": "s"
      },
      "D3": {
        "v": "Courier New",
        "s": {
          "font": {
            "name": "Courier New",
            "sz": 14
          }
        },
        "t": "s"
      },
      "A4": {
        "v": 0.618033989,
        "t": "n"
      },
      "B4": {
        "v": 0.618033989,
        "t": "n"
      },
      "C4": {
        "v": 0.618033989,
        "t": "n"
      },
      "D4": {
        "v": 0.618033989,
        "t": "n",
        "s": {
          "numFmt": "0.00%"
        }
      },
      "E4": {
        "v": 0.618033989,
        "t": "n",
        "s": {
          "numFmt": "0.00%",
          "fill": {
            "fgColor": {
              "rgb": "FFFFCC00"
            }
          }
        }
      },
      "A5": {
        "v": 0.618033989,
        "t": "n",
        "s": {
          "numFmt": "0%"
        }
      },
      "B5": {
        "v": 0.618033989,
        "t": "n",
        "s": {
          "numFmt": "0.0%"
        }
      },
      "C5": {
        "v": 0.618033989,
        "t": "n",
        "s": {
          "numFmt": "0.00%"
        }
      },
      "D5": {
        "v": 0.618033989,
        "t": "n",
        "s": {
          "numFmt": "0.000%"
        }
      },
      "E5": {
        "v": 0.618033989,
        "t": "n",
        "s": {
          "numFmt": "0.0000%"
        }
      },
      "F5": {
        "v": 0,
        "t": "n",
        "s": {
          "numFmt": "0.00%;\\(0.00%\\);\\-;@",
          "fill": {
            "fgColor": {
              "rgb": "FFFFCC00"
            }
          }
        }
      },
      "A6": {
        "v": "Sat Mar 21 2015 23:47:34 GMT-0400 (EDT)",
        "t": "s"
      },
      "B6": {
        "v": 42084.99137416667,
        "t": "n"
      },
      "C6": {
        "v": 42084.99137416667,
        "s": {
          "numFmt": "d-mmm-yy"
        },
        "t": "n"
      },
      "A7": {
        "v": "left",
        "s": {
          "alignment": {
            "horizontal": "left"
          }
        },
        "t": "s"
      },
      "B7": {
        "v": "center",
        "s": {
          "alignment": {
            "horizontal": "center"
          }
        },
        "t": "s"
      },
      "C7": {
        "v": "right",
        "s": {
          "alignment": {
            "horizontal": "right"
          }
        },
        "t": "s"
      },
      "A8": {
        "v": "vertical",
        "s": {
          "alignment": {
            "vertical": "top"
          }
        },
        "t": "s"
      },
      "B8": {
        "v": "vertical",
        "s": {
          "alignment": {
            "vertical": "center"
          }
        },
        "t": "s"
      },
      "C8": {
        "v": "vertical",
        "s": {
          "alignment": {
            "vertical": "bottom"
          }
        },
        "t": "s"
      },
      "A9": {
        "v": "indent",
        "s": {
          "alignment": {
            "indent": "1"
          }
        },
        "t": "s"
      },
      "B9": {
        "v": "indent",
        "s": {
          "alignment": {
            "indent": "2"
          }
        },
        "t": "s"
      },
      "C9": {
        "v": "indent",
        "s": {
          "alignment": {
            "indent": "3"
          }
        },
        "t": "s"
      },
      "A10": {
        "v": "In publishing and graphic design, lorem ipsum is a filler text commonly used to demonstrate the graphic elements of a document or visual presentation. ",
        "s": {
          "alignment": {
            "wrapText": 1,
            "horizontal": "right",
            "vertical": "center",
            "indent": 1
          }
        },
        "t": "s"
      },
      "A11": {
        "v": 41684.35264774306,
        "s": {
          "numFmt": "m/d/yy"
        },
        "t": "n"
      },
      "B11": {
        "v": 41684.35264774306,
        "s": {
          "numFmt": "d-mmm-yy"
        },
        "t": "n"
      },
      "C11": {
        "v": 41684.35264774306,
        "s": {
          "numFmt": "h:mm:ss AM/PM"
        },
        "t": "n"
      },
      "D11": {
        "v": 42084.99137416667,
        "s": {
          "numFmt": "m/d/yy"
        },
        "t": "n"
      },
      "E11": {
        "v": 42065.02247239584,
        "s": {
          "numFmt": "m/d/yy"
        },
        "t": "n"
      },
      "F11": {
        "v": 42084.99137416667,
        "s": {
          "numFmt": "m/d/yy h:mm:ss AM/PM"
        },
        "t": "n"
      },
      "A12": {
        "v": "Apple",
        "s": {
          "border": {
            "top": {
              "style": "thin"
            },
            "left": {
              "style": "thin"
            },
            "right": {
              "style": "thin"
            },
            "bottom": {
              "style": "thin"
            }
          }
        },
        "t": "s"
      },
      "C12": {
        "v": "Apple",
        "s": {
          "border": {
            "diagonalUp": 1,
            "diagonalDown": 1,
            "top": {
              "style": "dashed",
              "color": {
                "auto": 1
              }
            },
            "right": {
              "style": "medium",
              "color": {
                "theme": "5"
              }
            },
            "bottom": {
              "style": "hair",
              "color": {
                "theme": 5,
                "tint": "-0.3"
              }
            },
            "left": {
              "style": "thin",
              "color": {
                "rgb": "FFFFAA00"
              }
            },
            "diagonal": {
              "style": "dotted",
              "color": {
                "auto": 1
              }
            }
          }
        },
        "t": "s"
      },
      "E12": {
        "v": "Pear",
        "s": {
          "border": {
            "diagonalUp": 1,
            "diagonalDown": 1,
            "top": {
              "style": "dashed",
              "color": {
                "auto": 1
              }
            },
            "right": {
              "style": "dotted",
              "color": {
                "theme": "5"
              }
            },
            "bottom": {
              "style": "mediumDashed",
              "color": {
                "theme": 5,
                "tint": "-0.3"
              }
            },
            "left": {
              "style": "double",
              "color": {
                "rgb": "FFFFAA00"
              }
            },
            "diagonal": {
              "style": "hair",
              "color": {
                "auto": 1
              }
            }
          }
        },
        "t": "s"
      },
      "A13": {
        "v": "Up 90",
        "s": {
          "alignment": {
            "textRotation": 90
          }
        },
        "t": "s"
      },
      "B13": {
        "v": "Up 45",
        "s": {
          "alignment": {
            "textRotation": 45
          }
        },
        "t": "s"
      },
      "C13": {
        "v": "Horizontal",
        "s": {
          "alignment": {
            "textRotation": 0
          }
        },
        "t": "s"
      },
      "D13": {
        "v": "Down 45",
        "s": {
          "alignment": {
            "textRotation": 135
          }
        },
        "t": "s"
      },
      "E13": {
        "v": "Down 90",
        "s": {
          "alignment": {
            "textRotation": 180
          }
        },
        "t": "s"
      },
      "F13": {
        "v": "Vertical",
        "s": {
          "alignment": {
            "textRotation": 255
          }
        },
        "t": "s"
      },
      "A14": {
        "v": "Font color test",
        "s": {
          "font": {
            "color": {
              "rgb": "FFC6EFCE"
            }
          }
        },
        "t": "s"
      },
      "!ref": "A1:F14"
    }
  }
}
XLSX.writeFile(workbook, OUTFILE, { defaultCellStyle: defaultCellStyle });
console.log("open " + OUTFILE)

