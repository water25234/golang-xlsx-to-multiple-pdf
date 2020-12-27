package cmd

import (
	"fmt"
	"os"

	"github.com/spf13/cobra"
)

var (
	folder string

	readfile string

	protection bool

	sendMail bool

	rootCmd = &cobra.Command{
		Use:   "generatePDF",
		Short: "Generate multiple pdf for your xlsx.",
		RunE: func(cmd *cobra.Command, args []string) (err error) {
			r, err := checkFlags()
			if err != nil {
				errMsg(err)
				return err
			}

			err = r.GenerateMultiplePdf()
			if err != nil {
				return err
			}

			if sendMail == true {
				err = r.SendMail()
				if err != nil {
					return err
				}
			}

			return nil
		},
	}
)

type flags struct {
	folder string

	readfile string

	protection bool

	sendMail bool
}

func init() {
	rootCmd.PersistentFlags().StringVarP(&readfile, "readfile", "r", "", "read file content for pin code (default read pinCodeFile.txt)")
	rootCmd.PersistentFlags().StringVarP(&folder, "folder", "f", "", "create folder (default folder name file)")
	rootCmd.PersistentFlags().BoolVarP(&protection, "protection", "p", false, "make pdf protection (default protection false)")
	rootCmd.PersistentFlags().BoolVarP(&sendMail, "sendMail", "s", false, "send Mail (default false means not send mail)")
}

func checkFlags() (fs *flags, err error) {
	if len(readfile) == 0 {
		return nil, fmt.Errorf("flags readfile is empty")
	}

	if len(folder) == 0 {
		return nil, fmt.Errorf("flags folder is empty")
	}

	fs = &flags{
		readfile:   readfile,
		folder:     folder,
		protection: protection,
		sendMail:   sendMail,
	}

	return fs, nil
}

func errMsg(msg interface{}) {
	fmt.Println("Error:", msg)
}

// Execute command
func Execute() {
	if err := rootCmd.Execute(); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}
