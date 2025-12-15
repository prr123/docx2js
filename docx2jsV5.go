// program that converts docx files into js Dom
//
// V2 add numbering
// -- add border and margin to artdiv
// -- fix append to nested lists
// -- change output dir to test.imgerp.eu
//
// V3 read styles from numbering.xml
//
// V4
// 1. fix mixed numbering
// -- expand dbg
//
// V5
// - implement Bold Italic
//

package main

import (
	"fmt"
	"os"
	"log"
	"strings"

//	numLib "github.com/prr123/docxNumbering/numLib"
	numLib "goDemo/goDocx/docx2js/numLib"

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
    jsFilnam := "/home/peter/cloud/domains/test.imgerp.eu/js/" + inFil + ".js"
//    jsFilnam := "out/" + inFil + ".js"

    if dbg {
        fmt.Printf("input:   %s\n", inFilnam)
        fmt.Printf("js file: %s\n", jsFilnam)
        fmt.Printf("output:  %s\n", outFilnam)
    }

	// create map for lists
	OrdMarkMap := make(map[string]string)
	OrdMarkMap["decimal"] = "decimal"
	OrdMarkMap["lowerLetter"] = "lower-alpha"
	OrdMarkMap["lowerRoman"] = "lower-roman"

	UnOrdMarkMap := make(map[int]string)
	UnOrdMarkMap[0] = "disc"
	UnOrdMarkMap[1] = "circle"
	UnOrdMarkMap[2] = "square"
	UnOrdMarkMap[3] = "disc"
	UnOrdMarkMap[4] = "circle"
	UnOrdMarkMap[5] = "square"
	UnOrdMarkMap[6] = "disc"
	UnOrdMarkMap[7] = "circle"
	UnOrdMarkMap[8] = "square"

	pBold := false
	pItalic := false
	pLists := true
	elCnt :=0

	jsFil, err := os.Create(jsFilnam)
	if err != nil {log.Fatalf("error -- could not create js file: %v\n", err)}
	defer jsFil.Close()

	//create azul js file
	rdoc, err := godocx.OpenDocument(inFilnam)
	if err !=nil {log.Fatalf("error -- opening doc: %v\n", err)}

	// retrieve numbering file
	numObj, err := numLib.GetNumObj(rdoc)
    if err != nil {pLists = false}
//log.Fatalf("error -- can't get numbering file: %v\n", err)}

    ML, err := numObj.CreNList()
    if pLists && err != nil {log.Fatalf("error -- CreNList: %v\n", err)}
	if pLists && dbg {ML.PrintDocxList()}

	body := rdoc.Document.Body
	if body == nil {log.Fatalf("error -- no document body found!")}

	jsFil.WriteString("let jsdoc= {\n")

	jsFil.WriteString("render() {\n")
	jsFil.WriteString("  const artDiv = document.createElement('div');\n")
	jsFil.WriteString("  artDiv.style.margin='10px';\n")
	jsFil.WriteString("  artDiv.style.border='1px dashed purple';\n")

	jsFil.WriteString("  let txtel={};\n")
	jsFil.WriteString("  let eltxt={};\n")

	liLCnt := make([]int,128)
	licnt :=0
	runcount := 0
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
				// bold

				//style
				styl := prop.Style
//				outFil.WriteString("** Par Property Style: ")
				if styl != nil  {
					if dbg {fmt.Fprintf(jsFil,"// styl: %s\n", styl.Val)}
					// beginning of list
					if styl.Val == "ListParagraph" {
						if !pLists {log.Fatalf("error -- no numbering file!")}
						dl := ML.DLists[CnumId-1]
						if dbg {fmt.Fprintf(jsFil,"// list numId: %d abs: %d Ord: %t\n", CnumId, dl.AbId , dl.Ord)}

						if Cilvl>Ilvl {
							for il:=Ilvl +1; il< Cilvl+1; il++ {
								if dl.Ord {
									fmt.Fprintf(jsFil,"const list_%d_%d = document.createElement('ol');\n",il, licnt)
									mark:=OrdMarkMap[dl.Mark[il]]
									fmt.Fprintf(jsFil,"  list_%d_%d.style.listStyleType = '%s';\n",il, licnt, mark)
								} else {
									fmt.Fprintf(jsFil,"const list_%d_%d = document.createElement('ul');\n",il, licnt)
									mark:=UnOrdMarkMap[il]
									fmt.Fprintf(jsFil,"  list_%d_%d.style.listStyleType = '%s';\n",il, licnt, mark)
								}
								fmt.Fprintf(jsFil,"  list_%d_%d.style.listStylePosition = 'inside';\n",il, licnt)
								pad := 20 * il +20
								fmt.Fprintf(jsFil, "  list_%d_%d.style.paddingInlineStart = '%dpx';\n", il, licnt, pad)

								liLCnt[il]=licnt
								licnt++
							}
						}
						if Cilvl<Ilvl {
							for il:=Ilvl; il>Cilvl; il-- {
								fmt.Fprintf(jsFil,"  list_%d_%d.appendChild(list_%d_%d);\n", il-1, liLCnt[il-1], il, liLCnt[il])
							}
						}
						Ilvl = Cilvl
						preIsList = true
					}
					// end of list
					if preIsList && styl.Val != "ListParagraph" {
						for il:=Ilvl; il>0; il-- {
							fmt.Fprintf(jsFil,"  list_%d_%d.appendChild(list_%d_%d);\n", il-1, liLCnt[il-1], il, liLCnt[il])
						}
						fmt.Fprintf(jsFil,"  artDiv.appendChild(list_0_%d);\n", liLCnt[0])
						preIsList = false
					}
					switch styl.Val {
					// normal paragraph
					case "":
						jsFil.WriteString("  txtel = document.createElement('p');\n")

					case "Title":
						jsFil.WriteString("  txtel = document.createElement('h1');\n")

					case "Heading1":
						jsFil.WriteString("  txtel = document.createElement('h2');\n")

					case "Heading2":
						jsFil.WriteString("  txtel = document.createElement('h3');\n")

					case "Heading3":
						jsFil.WriteString("  txtel = document.createElement('h4');\n")

					case "ListParagraph":
						jsFil.WriteString(" txtel = document.createElement('li');\n")

					default:
						log.Fatalf("error -- unknown style: %s\n", styl.Val)
					} // end switch

				} else { // no styl
					if len(cpar.Children) == 0 {
						jsFil.WriteString("  txtel=document.createElement('br');\n")
					} else {
						jsFil.WriteString("  txtel = document.createElement('p');\n")
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
						fmt.Fprintf(jsFil,"  list_%d_%d.appendChild(list_%d_%d);\n", il-1, liLCnt[il-1], il, liLCnt[il])
					}
					fmt.Fprintf(jsFil,"  artDiv.appendChild(list_0_%d);\n", liLCnt[0])
					preIsList = false
					Ilvl = -1
				}
				if len(cpar.Children) >0 {
					jsFil.WriteString("  txtel = document.createElement('p');\n")
				} else {
					jsFil.WriteString("  brel = document.createElement('br');\n")
					jsFil.WriteString("  artDiv.appendChild(brel);\n")
					continue
				}
			}
			if dbg { fmt.Fprintf(jsFil, "// children: %d\n", len(cpar.Children))}
			count:=0
			for _, pch := range cpar.Children {
				count++
				if dbg {fmt.Fprintf(jsFil, "// child %d:\n", count)}
				if pch.Link != nil {
					if dbg {jsFil.WriteString("//   has Link\n")}
				}
				if pch.Run != nil {
					prop := pch.Run.Property
					if prop != nil {
                        if  prop.Bold != nil {
                            pBold = !pBold
                            if dbg {fmt.Fprintf(jsFil,"//     Bold %t\n", pBold)}
                        }
                        if prop.Italic != nil {
                            pItalic = !pItalic
                            if dbg {fmt.Fprintf(jsFil,"//     Italic %t\n", pItalic)}
                        }
                    }
					elCnt++
					fmt.Fprintf(jsFil,"  const spEl%d = document.createElement('span');\n",elCnt)
//					fmt.Fprintf(jsFil,"  spEl%d.textContent='%s';\n", runcount, rch.Text.Text)
					if !pBold {fmt.Fprintf(jsFil,"  spEl%d.style.fontWeight='normal';\n",elCnt)}
					if pBold {fmt.Fprintf(jsFil,"  spEl%d.style.fontWeight='bold';\n",elCnt)}
					if !pItalic {fmt.Fprintf(jsFil,"  spEl%d.style.fontStyle='normal';\n",elCnt)}
					if pItalic {fmt.Fprintf(jsFil,"  spEl%d.style.fontStyle='italic';\n",elCnt)}

					if dbg { fmt.Fprintf(jsFil, "// pch.run.children: %d\n", len(pch.Run.Children))}
					for  _, rch := range pch.Run.Children {
						runcount++
							//fmt.Fprintf(jsFil,"//   RunChild %d\n", runcount)
						if rch.Text == nil {continue}
                        fmt.Fprintf(jsFil,"  eltxt = document.createTextNode('%s');\n",rch.Text.Text)
                        fmt.Fprintf(jsFil,"  spEl%d.appendChild(eltxt);\n",elCnt)
					}
					fmt.Fprintf(jsFil,"  txtel.appendChild(spEl%d);\n", elCnt)
				} // run
			} //cpar.children
			if preIsList {
				fmt.Fprintf(jsFil,"  list_%d_%d.appendChild(txtel);\n",Ilvl, liLCnt[Ilvl])
			} else {
				fmt.Fprintf(jsFil,"  artDiv.appendChild(txtel);\n")
			}
		} //par
		if child.Table != nil {
			if dbg {jsFil.WriteString("// has table:\n")}
		}
	} // body


	if Ilvl> 0 {
		for il:=Ilvl; il>0; il-- {
			fmt.Fprintf(jsFil,"  list_%d_%d.appendChild(list_%d_%d);\n", il-1, liLCnt[il-1], il, liLCnt[il])
		}
		fmt.Fprintf(jsFil,"  artDiv.appendChild(list_0_%d);\n", liLCnt[0])
	}

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
               // NumProp
                if prop.NumProp != nil {
						outFil.WriteString("// ** Par Property NumPr: \n")
                    	Cilvl := prop.NumProp.ILvl.Val
                    	CnumId := prop.NumProp.NumID.Val
                    	fmt.Fprintf(outFil,"//lEV: %d cLevel: %d NumId: %d\n", Ilvl, Cilvl, CnumId)
           		}

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
				fmt.Fprintf(outFil,"children: %d\n", len(cpar.Children))
			}
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
