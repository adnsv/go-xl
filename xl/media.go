package xl

import (
	"hash/fnv"
)

func BlobHash(blob []byte) uint64 {
	h := fnv.New64()
	h.Write(blob)
	return h.Sum64()
}
