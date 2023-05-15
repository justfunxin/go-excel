package excel

import (
	"fmt"
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

type User struct {
	Id       int       `xlsx:"账号ID"`
	Name     string    `xlsx:"账号名"`
	Birthday time.Time `xlsx:"生日"`
	Interest []string  `xlsx:"兴趣"`
	Numbers  []int     `xlsx:"数字"`
}

func TestCreate(t *testing.T) {
	users := &[]User{
		{
			Id:       1,
			Name:     "Test1",
			Birthday: time.Now(),
			Interest: []string{"篮球", "户外"},
			Numbers:  []int{1, 2},
		},
		{
			Id:       2,
			Name:     "Test2",
			Birthday: time.Now(),
			Interest: []string{"篮球", "户外"},
			Numbers:  []int{1, 2},
		},
	}
	f, err := NewFile(users)
	if err != nil {
		fmt.Println(err)
		return
	}
	if err := f.SaveAs("Test1.xlsx"); err != nil {
		fmt.Println(err)
	}
}

func TestReadFromFile(t *testing.T) {
	users := &[]User{}
	err := GetRowsFromFile("Test1.xlsx", users)
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println(users)
}

func TestRead(t *testing.T) {
	f, err := excelize.OpenFile("Test1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	users := &[]User{}
	err = GetRowsFromFile("Test1.xlsx", users)
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println(users)
}
