package main

import (
	"fmt"
	"github.com/gin-gonic/gin"
	"log"
	"net/http"
)

func main() {
	router := gin.Default()
	// Set a lower memory limit for multipart forms (default is 32 MiB)
	router.MaxMultipartMemory = 8 << 20  // 8 MiB
	router.POST("/upload", func(c *gin.Context) {
		// Multipart form
		file, _ := c.FormFile("file")
		log.Println(file.Filename)
		c.SaveUploadedFile(file, "D:/auditReport/server/files")
		//files := form.File["upload[]"]
		//
		//for _, file := range files {
		//	log.Println(file.Filename)
		//
		//	// Upload the file to specific dst.
		//	c.SaveUploadedFile(file, "D:/auditReport/server/files")
		//}
		c.String(http.StatusOK, fmt.Sprintf("'%s' uploaded!", file.Filename))
	})
	router.Run(":8080")
}