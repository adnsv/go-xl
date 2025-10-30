package xl

import (
	"hash/fnv"
)

// BlobHash computes a 64-bit FNV-1 hash of the given byte slice.
// This is used to generate unique identifiers for embedded media files.
func BlobHash(blob []byte) uint64 {
	h := fnv.New64()
	h.Write(blob)
	return h.Sum64()
}
