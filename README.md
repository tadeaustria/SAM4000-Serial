# VBA Serial interface for SAM 4000

## Description

This VBA interface is designed to read results via serial interface from the SAM4000 device[^1].

![SAM4000](http://www.edelweiss-radling.de/images/Bilder_Homepage/schiesssport/sam4000.png)

This implementation offers an option to receive results of the machine and process them in office products (e.g. MS Excel, MS Access).

## Usage

### Setup

Include `modCOMM.bas`, `SAM4000.bas`, `Serie.cls` and `Shot.cls` into your project. Initially method `InitSAM` must be called to initialize the communication.

```
    Call InitSAM
```
Afterwards `GetSerie({amount})` can be called to receive {amount} many shot results from the machine. The function `GetSerie` will show infoboxes that guide through the process (currently only _German_ is supported).
```
    Dim someSerie As Serie
    Set someSerie = GetSerie(10)
```

### Mocking

If no machine is available, include `SAM4000Mock.bas` instead of `SAM4000.bas`[^2]. This module will include mocked versions of `InitSAM` and especially `GetSerie` that will return a simulated result with the same structure as the communication with the machine would have.

### Result

Via the `Serie` object, information about the whole series can be obtained, of with the `Shot` property all shots and so their data can be accessed via index.

![SerieShotUML](img/UML.svg)

## Credits

`modCOMM.bas` was originally developed by David M. Hitchner[^3], but modified to support VBA7 for modern systems.

[^1]: [RM-IV](https://www.disag.de/produkte/rm-iv/) may also be compatible, but was never tested.
[^2]: `modCOMM.bas` is not required either.
[^3]: [Serial Port Communication](http://www.thescarms.com/vbasic/CommIO.aspx)