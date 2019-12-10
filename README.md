# Mutify

A simple tool that mutes Windows when Spotify transmits advertisement and restore the volume afterwords.

## Requirements

- Windows XP, 7, 8, 10

## Installing

Download the version you need:

* [Windows 32bit](https://github.com/filippotoso/mutify/raw/master/bin/Win32/Release/Mutify.exe)
* [Windows 64bit](https://github.com/filippotoso/mutify/raw/master/bin/Win64/Release/Mutify.exe)

Save it wherever you want.

Create a shortcut to the executable and place it in the startup folder (Start => Programs => Startup folder).

In this way it will start automatically with your operating system.

## How does it work

When the application is running it checks for the Spotify app.
When it starts streaming advertising, it mutes the computer.
When the advertising ends, the volume is set to the previous level.