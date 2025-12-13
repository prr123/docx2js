package main

import (
	"fmt"
	"os"
	"log"
	"strings"
	"github.com/gomutex/godocx"

    util "github.com/prr123/utility/utilLib"
)

func main() {

    numarg := len(os.Args)
    flags:=[]string{"dbg", "in"}

    useStr := "/in=<infile> [/dbg]"
    helpStr := "doxc parsing program"

    if numarg > len(flags) +1 {
        fmt.Println("too many arguments in cl!")
        fmt.Println("usage: %s %s\n", os.Args[0], useStr)
        os.Exit(-1)
    }

    if numarg == 1 || (numarg > 1 && os.Args[1] == "help") {
        fmt.Printf("help: %s\n", helpStr)
        fmt.Printf("usage is: %s %s\n", os.Args[0], useStr)
        os.Exit(1)
    }

    flagMap, err := util.ParseFlags(os.Args, flags)
    if err != nil {log.Fatalf("util.ParseFlags: %v\n", err)}

    dbg:= false
    _, ok := flagMap["dbg"]
    if ok {dbg = true}

    inFil := ""
    inval, ok := flagMap["in"]
    if !ok {
        log.Fatalf("error -- no in flag provided!\n")
    } else {
        if inval.(string) == "none" {log.Fatalf("error -- no input file name provided!\n")}
        inFil = inval.(string)
		idx := strings.IndexByte(inFil, '.')
		if idx > -1 {log.Fatalf("error -- infile <%s> has an extension!\n", inFil)}
    }


    inFilnam := "/home/peter/go/src/goDemo/goDocx/docx/" + inFil + ".docx"
    outFilnam := "out/" + inFil + ".docxTxt"
    jsFilnam := "out/" + inFil + ".js"

    if dbg {
        fmt.Printf("input:  %s\n", inFilnam)
        fmt.Printf("output: %s\n", outFilnam)
    }

	jsFil, err := os.Create(jsFilnam)
	if err != nil {log.Fatalf("error -- could not create js file: %v\n", err)}
	defer jsFil.Close()

	//create azul js file
	rdoc, err := godocx.OpenDocument(inFilnam)
	if err !=nil {log.Fatalf("error -- opening doc: %v\n", err)}

//	root := &rdoc
//	fmt.Printf("root: %v\n", root)

	body := rdoc.Document.Body
	if body == nil {log.Fatalf("error -- no document body found!")}

	jsFil.WriteString("let jsdoc= {\n")
	jsFil.WriteString("render() {\n")
	jsFil.WriteString("  const artDiv = document.createElement('div');\n")

	chParCount := 0
	preIsList := false
	Ilvl := -1
	for _, child := range body.Children {

		if child.Para != nil {

			Cilvl := -1
			CnumId := -1
			chParCount++
//			outFil.WriteString("Para: \n")
			p:=child.Para
			// get Paragraph
			cpar := p.GetCT()
//			outFil.WriteString("** Par Property: \n")
			prop:= cpar.Property
			if prop != nil {
			// https://github.com/gomutex/godocx/blob/main/wml/ctypes/pPr.go#L12
				// NumProp
				if prop.NumProp != nil {
					if dbg {jsFil.WriteString("// ** Par Property NumPr: \n")}
					Cilvl = prop.NumProp.ILvl.Val
					CnumId = prop.NumProp.NumID.Val
					if dbg {fmt.Fprintf(jsFil,"//lEV: %d cLevel: %d NumId: %d\n", Ilvl, Cilvl, CnumId)}
				}
				//style
				styl := prop.Style
//				outFil.WriteString("** Par Property Style: ")
				if styl != nil  {
					if dbg {fmt.Fprintf(jsFil,"// styl: %s\n", styl.Val)}
					// beginning of list
					if styl.Val == "ListParagraph" {
						if Cilvl>Ilvl {
							for il:=Ilvl +1; il< Cilvl+1; il++ {
								fmt.Fprintf(jsFil," const list_%d = document.createElement('ol');\n",il)
							}
						}
						Ilvl = Cilvl
						preIsList = true
					}
					// end of list
					if preIsList && styl.Val != "ListParagraph" {
						for il:=Ilvl; il>0; il-- {
							fmt.Fprintf(jsFil,"  list_%d.appendChild(list_%d);\n", il-1, il)
						}
						fmt.Fprintf(jsFil,"  artDiv.appendChild(list_1);\n")
						preIsList = false
					}
					switch styl.Val {
					// normal paragraph
					case "":
						jsFil.WriteString("  const txtel = document.createElement('p');\n")

					case "Title":
						jsFil.WriteString("  const txtel = document.createElement('h1');\n")

					case "Heading1":
						jsFil.WriteString("  const txtel = document.createElement('h2');\n")

					case "Heading2":
						jsFil.WriteString("  const txtel = document.createElement('h3');\n")

					case "Heading3":
						jsFil.WriteString("  const txtel = document.createElement('h4');\n")

					case "ListParagraph":
						jsFil.WriteString(" const txtel = document.createElement('li');\n")

					default:
						log.Fatalf("error -- unknown style: %s\n", styl.Val)
					} // end switch

				} else { // no styl
					if len(cpar.Children) == 0 {
						jsFil.WriteString("  const txtel=document.createElement('br');\n")
					} else {
						jsFil.WriteString("  const txtel = document.createElement('p');\n")
					}
					//outFil.WriteString("no style")
				}

				//keepnext

				//Border

				//Justification

			} else  { // no prop
				// could be either an empty line or a simple paragraph
				if dbg {jsFil.WriteString("//  no prop\n")}
				if preIsList {
					for il:=Ilvl; il>0; il-- {
						fmt.Fprintf(jsFil,"  list_%d.appendChild(list_%d);\n", il-1, il)
					}
					fmt.Fprintf(jsFil,"  artDiv.appendChild(list_0);\n")
					preIsList = false
					Ilvl = -1
				}
				if len(cpar.Children) >0 {
					jsFil.WriteString("  const txtel = document.createElement('p');\n")
				} else {
					jsFil.WriteString("  const brel = document.createElement('br');\n")
					jsFil.WriteString("  artDiv.appendChild(brel);\n")
					continue
				}
			}
			//outFil.WriteString("children: \n")
			count:= 0
			for _, pch := range cpar.Children {
				count++
				if dbg {fmt.Fprintf(jsFil, "// child %d:\n", count)}
				if pch.Link != nil {
					if dbg {jsFil.WriteString("//   has Link\n")}
				}
				if pch.Run != nil {
					//outFil.WriteString("  Run Els\n")
					runcount := 0
					for  _, rch := range pch.Run.Children {
						runcount++
						//fmt.Fprintf(outFil,"    RunChild %d\n", runcount)
						if rch.Text != nil {
							if dbg {fmt.Fprintf(jsFil,"// runcount: %d\n", runcount)}
							fmt.Fprintf(jsFil,"  const eltxt = document.createTextNode('%s');\n",rch.Text.Text)
							fmt.Fprintf(jsFil,"  txtel.appendChild(eltxt);\n")
						}
					}
				} // run
			} //cpar.children
			if preIsList {
				fmt.Fprintf(jsFil,"  list_%d.appendChild(txtel);\n",Ilvl)
			} else {
				fmt.Fprintf(jsFil,"  artDiv.appendChild(txtel);\n")
			}
		} //par
		if child.Table != nil {
			if dbg {jsFil.WriteString("// has table:\n")}
		}
	} // body


	for il:=Ilvl; il>0; il-- {
		fmt.Fprintf(jsFil,"  list_%d.appendChild(list_%d);\n", il-1, il)
	}
	fmt.Fprintf(jsFil,"  artDiv.appendChild(list_0);\n")

	jsFil.WriteString("  return artDiv;\n},\n")
	jsFil.WriteString("};\n")

	if !dbg {
		log.Printf("*** success  ***\n")
		os.Exit(0)
	}

	outFil, err := os.Create(outFilnam)
	if err != nil {log.Fatalf("error -- could not create out file: %v\n", err)}
	defer outFil.Close()

//	rdoc, err := godocx.OpenDocument(inFilnam)
//	if err !=nil {log.Fatalf("error -- opening doc: %v\n", err)}

//	body := rdoc.Document.Body

	for _, child := range body.Children {
		if child.Para != nil {
			outFil.WriteString("Para: \n")
			p:=child.Para
			// get Paragraph
			cpar := p.GetCT()
			outFil.WriteString("** Par Property: \n")
			prop:= cpar.Property
			if prop != nil {
			// https://github.com/gomutex/godocx/blob/main/wml/ctypes/pPr.go#L12
				//style
				styl := prop.Style
				outFil.WriteString("** Par Property Style: ")
				if styl != nil  {
					outFil.WriteString(styl.Val)
				} else {
					outFil.WriteString("no style")
				}
				outFil.WriteString("\n")

				//keepnext

				//Border

				//Justification

			} else  {
				outFil.WriteString("  no prop\n")
			}
			outFil.WriteString("children: \n")
			count:= 0
			for _, pch := range cpar.Children {
				count++
				fmt.Fprintf(outFil, "child %d:\n", count)
				if pch.Link != nil {
					outFil.WriteString("   has Link\n")
				}
				if pch.Run != nil {
					outFil.WriteString("  Run Els\n")
					runcount := 0
					for  _, rch := range pch.Run.Children {
						runcount++
						fmt.Fprintf(outFil,"    RunChild %d\n", runcount)
						if rch.Text != nil {
							fmt.Fprintf(outFil,"> %s <\n", rch.Text.Text)
						}
					}
				}
			}
		}
		if child.Table != nil {
			outFil.WriteString("has table:\n")
		}
	}
	log.Printf("*** success with debug ***\n")
}
