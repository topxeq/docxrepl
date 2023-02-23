package docxrepl

import (
	"bytes"
	"fmt"
)

func ConvertToReplaceMap(vA interface{}) interface{} {
	switch nv := vA.(type) {
	case map[string]interface{}:
		return PlaceholderMap(nv)
	case map[string]string:
		rs := PlaceholderMap{}

		for k, v := range nv {
			rs[k] = v
		}
		return rs
	case []string:
		rs := PlaceholderMap{}

		lenT := len(nv) / 2
		for i := 0; i < lenT; i++ {
			rs[nv[i*2]] = nv[i*2+1]
		}

		return rs
	case []interface{}:
		rs := PlaceholderMap{}

		lenT := len(nv) / 2
		for i := 0; i < lenT; i++ {
			rs[fmt.Sprintf("%v", nv[i*2])] = nv[i*2+1]
		}

		return rs
	}

	return fmt.Errorf("failed to convert map: %T(%v)", vA, vA)
}

func ReplaceInWordBytes(vA []byte, mapA interface{}) interface{} {
	if vA == nil {
		return nil
	}

	replaceMap := ConvertToReplaceMap(mapA)

	if nv1, ok := replaceMap.(error); ok {
		return nv1
	}

	doc, err := OpenBytes(vA)
	if err != nil {
		return err
	}

	err = doc.ReplaceAll(replaceMap.(PlaceholderMap))
	if err != nil {
		return err
	}

	var bufT bytes.Buffer

	err = doc.Write(&bufT)

	if err != nil {
		return err
	}

	return bufT.Bytes()
}
