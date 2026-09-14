#!/bin/zsh
cd "${0:A:h}"
if [[ -s "$HOME/.nvm/nvm.sh" ]]; then source "$HOME/.nvm/nvm.sh"; fi
npm start
read '?按回车关闭窗口…'
