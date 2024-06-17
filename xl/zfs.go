package xl

import (
	"archive/zip"
	"io"
	"os"
	"path/filepath"
	"strings"
)

type Storage interface {
	WriteBlob(path string, blob []byte) error
}

type DirStorage struct {
	Dir string
}

type ZipStorage struct {
	z *zip.Writer
}

func NewDirStorage(dir string) *DirStorage {
	return &DirStorage{
		Dir: dir,
	}
}

func (ds *DirStorage) WriteBlob(path string, blob []byte) error {
	path = strings.TrimPrefix(path, "/")
	fn := filepath.Join(ds.Dir, path)
	err := os.MkdirAll(filepath.Dir(fn), 0777)
	if err != nil {
		return err
	}
	return os.WriteFile(fn, blob, 0666)
}

func NewZipStorage(out io.Writer) *ZipStorage {
	return &ZipStorage{z: zip.NewWriter(out)}
}

func (zs *ZipStorage) WriteBlob(path string, blob []byte) error {
	path = strings.TrimPrefix(path, "/")
	f, err := zs.z.Create(path)
	if err != nil {
		return err
	}
	_, err = f.Write(blob)
	return err
}

func (zs *ZipStorage) Close() {
	zs.z.Close()
}
