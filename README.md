# Pingoji
Light VBS script to visualize ping results in small IE window floating always on top

Ever being in a train, working on your PC and loosing network connection?

Pingoji helps you immediatelly recognize what's the situation with the network coverage.

It provides visualization of network availability based on built-in ping of Windows system.

Prerequisites: the AutoHotkey app is needed to make the IE window floating on top

No install is needed - just download the 2 files:
  * Pingoji.vbs - the main program - it runs the ping command for network monitoring and IE for visualization
  * SetAlwaysOnTop.ahk - force the IE window float on top of the screen

The script is pinging 8.8.8.8 host (Google's DNS) by default, you can change it by editing the source code of VBS script.

When it runs normally it shows the connection as stable - timing (in miliseconds) for the last 5 pings is shown as green blocks.
![image](https://github.com/user-attachments/assets/eab83d84-89c8-4c2e-92fa-be945d5efeae)

The connection is considered stable in case 5 consequitive ping request successfully returned a response.

The darker green a block is the better connection you get.

![image](https://github.com/user-attachments/assets/a5763b71-8091-4338-9c6d-43c70fd397fb)


When you're loosing the connection the blocks are getting red and timing is indicated as "x" (as uknown).

![image](https://github.com/user-attachments/assets/e28a82c6-37b4-4391-9dfa-5a6f5d1b8c0e)

