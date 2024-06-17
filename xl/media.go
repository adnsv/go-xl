package xl

import (
	"hash/fnv"

	"github.com/google/uuid"
)

func BlobHash(blob []byte) uuid.UUID {
	h := fnv.New128()
	h.Write(blob)
	uid, _ := uuid.FromBytes(h.Sum([]byte{}))
	return uid
}
