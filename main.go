package main // import "hello-mdb"

import (
	"database/sql"
	"fmt"
	"os"
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	_ "github.com/mattn/go-adodb"
)

var provider string

func createMdb(f string) error {
	unk, err := oleutil.CreateObject("ADOX.Catalog")
	if err != nil {
		return err
	}
	defer unk.Release()
	cat, err := unk.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return err
	}
	defer cat.Release()
	provider = "Microsoft.Jet.OLEDB.4.0"
	r, err := oleutil.CallMethod(cat, "Create", "Provider="+provider+";Data Source="+f+";")
	if err != nil {
		provider = "Microsoft.ACE.OLEDB.12.0"
		r, err = oleutil.CallMethod(cat, "Create", "Provider="+provider+";Data Source="+f+";")
		if err != nil {
			return err
		}
	}
	r.Clear()
	return nil
}

func main() {
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	f := "./example.mdb"

	os.Remove(f)

	err := createMdb(f)
	if err != nil {
		fmt.Println("create mdb", err)
		return
	}

	db, err := sql.Open("adodb", "Provider="+provider+";Data Source="+f+";")
	if err != nil {
		fmt.Println("open", err)
		return
	}
	defer db.Close()

	sql := ""

	sql = `
	CREATE TABLE foo (
		id INT NOT NULL PRIMARY KEY,
		name TEXT NOT NULL,
		created DATETIME NOT NULL
	)`

	_, err = db.Exec(sql)
	if err != nil {
		fmt.Println("CREATE TABLE", err)
		return
	}

	tx, err := db.Begin()
	if err != nil {
		fmt.Println(err)
		return
	}

	sql = `INSERT INTO foo (
		id, name, created
	) VALUES(
		?, ?, ?
	)`
	stmt, err := tx.Prepare(sql)
	if err != nil {
		fmt.Println("INSERT", err)
		return
	}
	defer stmt.Close()

	for idx := 0; idx < 1000; idx++ {
		name := fmt.Sprintf("안녕세상 - %03d", idx)
		create := time.Now()

		_, err = stmt.Exec(idx, name, create)
		if err != nil {
			fmt.Println("EXEC", err)
			return
		}
	}
	tx.Commit()

	sql = `SELECT
		id, name, created
	FROM foo`
	rows, err := db.Query(sql)
	if err != nil {
		fmt.Println("select", err)
		return
	}
	defer rows.Close()

	for rows.Next() {
		var id int
		var name string
		var created time.Time
		err = rows.Scan(&id, &name, &created)
		if err != nil {
			fmt.Println("SCAN", err)
			return
		}
		fmt.Println(id, name, created)
	}
}
