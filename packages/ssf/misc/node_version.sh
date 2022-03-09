#! /usr/bin/env bash

# This script will check the current version of node and install another version
# of npm if node is version 0.8

version=$(node --version)

if [[ $version =~ v0\.8\. ]]
then
  npm install -g npm@4.3.0
fi