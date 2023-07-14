#!/usr/bin/bash

tmpFILE="/tmp/pptx.refresh"

job(){
  clear
  echo python3 bard.py
  python3 ./bard.py
}

main(){

  trap 'exit 0;' INT

  while true;do
    sleep 1
    echo
    while [ ! -f "$tmpFILE" ];do
      sleep 1
      echo -n "."
    done
    job
    rm -f "$tmpFILE"
  done
}


main "$@"
