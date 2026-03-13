package main

import (
	"fmt"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"text/template"
	"time"

	docx "github.com/lukasjarosch/go-docx"
)

const pageHTML = `
<!doctype html>
<html lang="ru">
<head>
  <meta charset="UTF-8">
  <title>Генератор документов</title>
  <style>
    body {
      font-family: Arial, sans-serif;
      max-width: 700px;
      margin: 40px auto;
      padding: 20px;
    }

    h1 {
      margin-bottom: 20px;
    }

    .field {
      margin-bottom: 16px;
    }

    label {
      display: block;
      margin-bottom: 6px;
      font-weight: bold;
    }

    input {
      width: 100%;
      padding: 10px;
      border: 1px solid #0e3fa8;
      border-radius: 8px;
      box-sizing: border-box;
    }

    button {
      padding: 12px 18px;
      border: none;
      border-radius: 8px;
      cursor: pointer;
    }
  </style>
</head>
<body>
  <h2>Заполнение шаблона</h2>

  <form method="post" action="/generate">
  <div class="field">
    <label>ФИО:</label><br>
	<input type="text" name="fio" style="width: 400px;" required>
  </div>
  <br>

  <div class="field">
    <label>ИНН:</label><br>
	<input type="text" name="inn" style="width: 400px;">
  </div>
  <br>

  <div class="field">
    <label>Расчётный счёт:</label><br>
	<input type="text" name="account" style="width: 400px;">
  </div>
  <br>

  <div class="field">
    <label>Банк:</label><br>
	<input type="text" name="bank" style="width: 400px;">
  </div>
  <br>

  <div class="field">
    <label>БИК:</label><br>
	<input type="text" name="bik" style="width: 400px;">
  </div>
  <br>

	<button type="submit">Сохранить</button>
	</form>
</body>
</html>
`

var tmpl = template.Must(template.New("page").Parse(pageHTML))

func main() {
	http.HandleFunc("/", formHandler)
	http.HandleFunc("/generate", generateHandler)

	log.Println("Server starts")
	log.Fatal(http.ListenAndServe(":8080", nil))
}

func formHandler(w http.ResponseWriter, r *http.Request) {
	if err := tmpl.Execute(w, nil); err != nil {
		http.Error(w, "Form error", http.StatusInternalServerError)
		return
	}
}

func generateHandler(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "invalid method (post only)", http.StatusMethodNotAllowed)
		return
	}

	if err := r.ParseForm(); err != nil {
		http.Error(w, "can not read file", http.StatusBadRequest)
	}

	fio := r.FormValue("fio")
	inn := r.FormValue("inn")
	account := r.FormValue("account")
	bank := r.FormValue("bank")
	bik := r.FormValue("bik")

	doc, err := docx.Open("Testfile.docx")
	if err != nil {
		http.Error(w, "can not open .docx file: "+err.Error(), http.StatusInternalServerError)
		return
	}

	replaceMap := docx.PlaceholderMap{
		"fio":     fio,
		"inn":     inn,
		"account": account,
		"bank":    bank,
		"bik":     bik,
	}

	if err := doc.ReplaceAll(replaceMap); err != nil {
		http.Error(w, "can not insert data: "+err.Error(), http.StatusInternalServerError)
		return
	}

	outFile := filepath.Join(os.TempDir(), fmt.Sprintf("result_%d.docx", time.Now().UnixNano()))
	if err := doc.WriteToFile(outFile); err != nil {
		http.Error(w, "can not save file: "+err.Error(), http.StatusInternalServerError)
		return
	}
	defer os.Remove(outFile)

	data, err := os.ReadFile(outFile)
	if err != nil {
		http.Error(w, "can not read a result"+err.Error(), http.StatusInternalServerError)
		return
	}

	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
	w.Header().Set("Content-Disposition", `attachment; filename="result.docx"`)
	w.WriteHeader(http.StatusOK)
	_, _ = w.Write(data)
}
