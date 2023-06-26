package main

import (
	"encoding/xml"
	"fmt"
	"io/ioutil"
	"log"
	"strings"
	"time"
	"github.com/xuri/excelize/v2"
	"golang.org/x/net/html"

)

type shop struct {
	Name       string      `xml:"name"`
	Company    string      `xml:"company"`
	URL        string      `xml:"url"`
	Currencies []Currency  `xml:"currencies>currency"`
	Categories []Category  `xml:"categories>category"`
	Offers     []Offer     `xml:"offers>offer"`
}

type Currency struct {
	ID   string `xml:"id,attr"`
	Rate string `xml:"rate,attr"`
}

type Category struct {
	ID       string `xml:"id,attr"`
	ParentID string `xml:"parentId,attr,omitempty"`
	Name     string `xml:",chardata"`
}

type Offer struct {
	ID          string    `xml:"id,attr"`
	Available   bool      `xml:"available,attr"`
	URL         string    `xml:"url"`
	Price       string    `xml:"price"`
	CurrencyID  string    `xml:"currencyId"`
	CategoryID  string    `xml:"categoryId"`
	Pictures    []Picture `xml:"picture"`
	Pickup      bool      `xml:"pickup"`
	Delivery    bool      `xml:"delivery"`
	Name        string    `xml:"name"`
	NameUkr     string    `xml:"name_ua"`
	Vendor      string    `xml:"vendor"`
	VendorCode  string    `xml:"vendorCode"`

	Description string `xml:"description"`
	DescrUkr    string `xml:"description_ua"`
}

type Picture struct {
	Value string `xml:",chardata"`
}

// CDATAText represents a text value wrapped within a CDATA section.
type CDATAText struct {
	Value string `xml:",cdata"`
}

// CloseHTMLTags закрывает незакрытые HTML-теги в строке
func CloseHTMLTags(input string) string {
	tokenizer := html.NewTokenizer(strings.NewReader(input))
	stack := make([]string, 0)

	for {
		tokenType := tokenizer.Next()
		if tokenType == html.ErrorToken {
			break
		}

		token := tokenizer.Token()
		switch tokenType {
		case html.StartTagToken:
			stack = append(stack, token.Data)
		case html.EndTagToken:
			if len(stack) > 0 {
				stack = stack[:len(stack)-1]
			}
		}
	}

	for len(stack) > 0 {
		tag := stack[len(stack)-1]
		input += fmt.Sprintf("</%s>", tag)
		stack = stack[:len(stack)-1]
	}

	return input
}

func main() {
	// Открытие файла Excel
	f, err := excelize.OpenFile("products.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	// Чтение данных из первой страницы Excel
	rows, err := f.GetRows("Export Products Sheet")
	if err != nil {
		log.Fatal(err)
	}

	// Создание структуры магазина
	shop := shop{
		Name:    "Allegro *UA*",
		Company: "Allegro *UA*",
		URL:     "Allegro *UA*",
		Currencies: []Currency{
			{ID: "USD", Rate: "CB"},
			{ID: "PLN", Rate: "1"},
			{ID: "BYN", Rate: "CB"},
			{ID: "KZT", Rate: "CB"},
			{ID: "EUR", Rate: "CB"},
		},
		Categories: []Category{},
		Offers:     []Offer{},
	}

	// Проход по строкам Excel
	for _, row := range rows[1:] {
		// Создание оффера
		offer := Offer{
			ID:         row[0],
			Available:  true,
			URL:        row[1],
			Price:      row[11],
			CurrencyID: "PLN",
			CategoryID: row[30],
			Pictures:   []Picture{},
			Pickup:     false,
			Delivery:   true,
			Name:       row[3],
			NameUkr:    row[4],
			Vendor:     "",
			VendorCode: row[27],
			Description: row[8],
			DescrUkr: row[9],
		
		}
		
		
		offer.Description = strings.Replace(offer.Description, `</td>`, ``, -1)
		offer.DescrUkr = strings.Replace(offer.DescrUkr, `</td>`, ``, -1)

		offer.DescrUkr = strings.ReplaceAll(offer.DescrUkr, "&#xA;", " ")
		offer.Description = strings.ReplaceAll(offer.Description, "&#xA;", " ")



offer.Description  = html.UnescapeString(offer.Description )

		// Добавление изображений
		pictures := row[17]
		if pictures != "" {
			imageUrls := strings.Split(pictures, ",")
			for _, imageUrl := range imageUrls {
				picture := Picture{Value: imageUrl}
				offer.Pictures = append(offer.Pictures, picture)
			}
		}

		// Добавление оффера в список офферов
		shop.Offers = append(shop.Offers, offer)
	}
	
	
	// Чтение данных из второй страницы Excel
	rows, err = f.GetRows("Export Groups Sheet")
	if err != nil {
		log.Fatal(err)
	}

	// Проход по строкам Excel
	for _, row := range rows[1:] {
		// Создание категории
		category := Category{
			ID:   row[4],
			Name: row[3],
		}

		// Добавление родительской категории, если есть
		if len(row) > 6 && row[6] != "" {
			category.ParentID = row[6]
		}

		// Добавление категории в список категорий
		shop.Categories = append(shop.Categories, category)
	}

	// Генерация XML
		output, err := xml.MarshalIndent(shop, "", "  ")
	if err != nil {
		log.Fatal(err)
	}
	
	xmlString := string(output)
	xmlString = strings.ReplaceAll(xmlString, "&#xA;", "\n")

	// Запись XML-файла
	header := `<?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE yml_catalog>
`

	date := time.Now().Format("2006-01-02 15:04")

	footer := "\n</yml_catalog>"

	xmlData := []byte(header + fmt.Sprintf("<yml_catalog date=\"%s\">\n", date) + string(output) + footer)

	err = ioutil.WriteFile("output.xml", xmlData, 0644)
	if err != nil {
		log.Fatal(err)
	}

	fmt.Println("Done!")
}
