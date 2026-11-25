# OpenVBS

**An open-source, VBScript-compatible BASIC scripting engine for Windows, Linux, and macOS.**

OpenVBS helps keep existing VBScript code usable while extending BASIC with modern language features. It is powered by the original **Jujube** scripting core.

**Looking for a VBScript alternative on Windows?** See [OpenVBS-HotPatch](https://github.com/yrm006/openvbs-hotpatch), a Windows Active Scripting (`IActiveScript`) implementation with examples for Windows Script Host, embedded COM hosts, and registration-free loading.

## Description

OpenVBS aims for VBScript compatibility and adds modern BASIC extensions. A Classic ASP layer is under development to support legacy website migration, including Linux deployments. Compatibility varies by host and feature; this project does not claim that every VBScript or Classic ASP application works unchanged.

## Sample.obs
Bingo Machine:

    option explicit: option json: randomize:

    dim machine = BingoMachine({})
    machine.shuffle(1000)

    dim taker = BingoTaker({}, machine)
    doeach(taker.take, function(v) wscript.echo(v): end function)



    function BingoMachine(self)
        self.balls   = [ 1, 2, 3, 4, 5, 6, 7, 8, 9,10,11,12,13,14,15,16,17,18,19,20,
                        21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,
                        41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,
                        61,62,63,64,65,66,67,68,69,70,71,72,73,74,75]
        self.shuffle = function(n)
                        dim last = self.balls.length-1
                        do while n--
                            dim r = int(rnd()*last)
                            dim tmp = self.balls(r)
                            self.balls(r) = self.balls(last)
                            self.balls(last) = tmp
                        loop
                    end function
        return self
    end function

    function BingoTaker(self, machine)
        self.machine = machine
        self.i       = machine.balls.length
        self.take    = function()
                        return if self.i then self.machine.balls(--self.i) else empty end if
                    end function
        return self
    end function

    function doeach(eachee, dofunc)
        10 DIM E = EACHEE()
        20 IF ISEMPTY(E) THEN EXIT FUNCTION: END IF
        30 DOFUNC(E)
        40 GOTO 10
    end function

## How to build

[Windows] x64/x86 Native Tools Command Prompt for VS 2022

    > nmake -f makefile.win

[OSX] Xcode gcc

    $ make -f makefile.osx

[Linux] GNU gcc

    $ make -f makefile.linux

## Usage
    $ ./oscript sample.obs
    $ echo "wscript.echo 123" | ./oscript

[Interactive Mode]

    $ ./oscript -i
    oscript| dim msg="hello"
    oscript| wscript.echo msg
    hello
    oscript| 

[more...](https://github.com/yrm006/openvbs/blob/master/readme.txt)

## Licence
MIT

## Author
[NaturalStyle PREMIUM](https://p.na-s.jp)

---
yrm.20261125
