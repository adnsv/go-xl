package xl

import (
	"archive/zip"
	"io"
	"os"
	"path/filepath"
	"strings"
)

// Storage is the interface for writing Excel file parts (XML and media files).
// Implementations can write to ZIP archives or directory structures.
type Storage interface {
	WriteBlob(path string, blob []byte) error
}

// DirStorage writes Excel file parts to a directory structure on disk.
// This is useful for debugging as it allows inspection of generated XML files.
type DirStorage struct {
	Dir string // Root directory path
}

// ZipStorage writes Excel file parts to a ZIP archive, creating a standard .xlsx file.
type ZipStorage struct {
	z *zip.Writer
}

// NewDirStorage creates a new directory-based storage that writes files to the specified directory.
// The directory will be created if it doesn't exist.
func NewDirStorage(dir string) *DirStorage {
	return &DirStorage{
		Dir: dir,
	}
}

// WriteBlob writes a file part to the directory structure.
// Creates any necessary parent directories automatically.
func (ds *DirStorage) WriteBlob(path string, blob []byte) error {
	path = strings.TrimPrefix(path, "/")
	fn := filepath.Join(ds.Dir, path)
	err := os.MkdirAll(filepath.Dir(fn), 0777)
	if err != nil {
		return err
	}
	return os.WriteFile(fn, blob, 0666)
}

// NewZipStorage creates a new ZIP-based storage that writes to the given writer.
// The writer is typically a file opened for writing (e.g., os.Create("output.xlsx")).
func NewZipStorage(out io.Writer) *ZipStorage {
	return &ZipStorage{z: zip.NewWriter(out)}
}

// WriteBlob writes a file part to the ZIP archive.
// Each part becomes a file entry in the ZIP with the specified path.
func (zs *ZipStorage) WriteBlob(path string, blob []byte) error {
	path = strings.TrimPrefix(path, "/")
	f, err := zs.z.Create(path)
	if err != nil {
		return err
	}
	_, err = f.Write(blob)
	return err
}

// Close finalizes the ZIP archive. Must be called after all writes are complete.
// Failure to call Close will result in an invalid/corrupted Excel file.
func (zs *ZipStorage) Close() {
	zs.z.Close()
}
