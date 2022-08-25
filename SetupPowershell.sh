#! /bin/zsh

set -e

# shellcheck disable=SC2154
trap 'ret=$?; test $ret -ne 0 && printf "failed\n\n" >&2; exit $ret' EXIT

fancy_echo() {
  local fmt="$1"; shift

  # shellcheck disable=SC2059
  printf "\n$fmt\n" "$@"
}

append_to_zshrc() {
  local text="$1" zshrc
  local skip_new_line="${2:-0}"

  if [ -w "$HOME/.zshrc.local" ]; then
    zshrc="$HOME/.zshrc.local"
  else
    zshrc="$HOME/.zshrc"
  fi

  if ! grep -Fqs "$text" "$zshrc"; then
    if [ "$skip_new_line" -eq 1 ]; then
      printf "%s\n" "$text" >> "$zshrc"
    else
      printf "\n%s\n" "$text" >> "$zshrc"
    fi
  fi
}

#Check if brew is installed, install if it is not
if ! command -v brew >/dev/null; then
  fancy_echo "Installing Homebrew ..."
  /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
  append_to_zshrc '# recommended by brew doctor'
  # shellcheck disable=SC2016
  append_to_zshrc 'export PATH="/usr/local/bin:$PATH"' 1
  export PATH="/usr/local/bin:$PATH"
  echo 'eval "$(/opt/homebrew/bin/brew shellenv)"' >> ~/.zprofile
  eval "$(/opt/homebrew/bin/brew shellenv)"
fi

#Check if rosetta 2 is installed (required for powershell)
if [[ \"`pkgutil --files com.apple.pkg.RosettaUpdateAuto`\" == \"\" ]]
  then 
    fancy_echo 'Rosetta not detected, installing...'
    sudo softwareupdate --install-rosetta --agree-to-license
  else
    fancy_echo 'Rosetta detected, skipping installation...'
fi 

brew install powershell
brew install openssl
pwsh -Command 'Install-Module -Name PowerShellGet'
pwsh -Command 'Install-Module -Name PSWSMan'
sudo pwsh -Command 'Install-WSMan'
pwsh -Command 'Install-Module -Name ExchangeOnlineManagement -AllowPrerelease'

echo "Done! You can now use powershell.  Just type 'pwsh' and get shellin' power!"