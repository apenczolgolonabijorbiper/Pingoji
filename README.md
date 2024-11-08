# &#8584; Pingoji - network monitoring tool
Light VBS script to visualize ping results in small IE window floating always on top - network monitoring tool.

Ever being in a train, working on your PC and loosing network connection?

Pingoji helps you immediatelly recognize what's the situation with the network coverage.

It provides visualization of network availability based on built-in ping of Windows system.

Prerequisites: the AutoHotkey app is needed to make the IE window floating on top

In case you don't want to install AutoHotKey (and have no Always On Top feature) just remove (or comment-out) the following line in the VBS script:

> objShell.Run "SetAlwaysOnTop.ahk", 0, False 

The script will run further but the IE window will behave normally as any other window.

No install is needed - just download the 2 files:
  * Pingoji.vbs - the main program - it runs the ping command for network monitoring and IE for visualization
  * SetAlwaysOnTop.ahk - force the IE window float on top of the screen

The script is pinging 8.8.8.8 host (Google's DNS) by default, you can change it by editing the source code of VBS script.

When it runs normally it shows the connection as stable - timing (in miliseconds) for the last 5 pings is shown as green blocks.

![image](https://github.com/user-attachments/assets/b6623acb-286e-4cee-ac6e-2c3d56806d4c)

The connection is considered stable in case 5 consequitive ping requests successfully returned their responses, otherwise the window signals a problem.

The darker green a block is the better connection you get. The blocks are moving from left to right.

![image](https://github.com/user-attachments/assets/91e5d33e-26f7-4b1b-97dc-acbee03a3ce4)

When you're loosing the connection the blocks are getting red and timing is indicated as "x" (as uknown).

![image](https://github.com/user-attachments/assets/e28a82c6-37b4-4391-9dfa-5a6f5d1b8c0e)

When run there is a "black window" open together with the IE - don't bother it, but let it run in the background (it just executes the ping).

![image](https://github.com/user-attachments/assets/8171971e-8940-4483-aee5-bc531b643952)

