package main

import (
	"log"

	"github.com/joho/godotenv"

	"github.com/water25234/golang-xlsx-to-multiple-pdf/cmd"
)

func main() {
	err := godotenv.Load()
	if err != nil {
		log.Fatalf("Error getting env, %v", err)
	}
	cmd.Execute()
}
