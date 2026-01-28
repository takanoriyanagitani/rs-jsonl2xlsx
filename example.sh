#!/bin/sh

oxname="./sample.d/out.xlsx"
ijson="./sample.d/input.jsonl"

mkdir -p ./sample.d

geninput(){
    jq -c -n '[
        {id: 1, category:"ui", input_setup:"launch", input_task:"to home", expected:"home", edge_result:"pass", edge_date:"2026-01-27", edge_tester:"JD"},
        {id: 2, category:"ui", input_setup:"stop",   input_task:"to home", expected:"500",  edge_result:"pass", edge_date:"2026-01-27", edge_tester:"JD"}
    ]' |
        jq -c '.[]' |
        dd if=/dev/stdin of="${ijson}" bs=1048576 status=none
}

test -f "${ijson}" || geninput

cat "${ijson}" |
    wazero \
        run \
        ./rs-jsonl2xlsx.wasm \
        -- \
        --sheet-name=Sheet3 |
    dd if=/dev/stdin of="${oxname}" bs=1048576 status=none
