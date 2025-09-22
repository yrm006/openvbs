OpenVBS
====

The Script Engine for the BASIC-2020.

## Description
This is a script engine for the modern BASIC powered by the Jujube (self made) core. It is compatible with VBScript and has the modern features added. In addition, the Classic-ASP layer is implemented(on the way) as a library, so you can port your website to the Linux. [more...](https://p.na-s.jp/openvbs.html)

## Sample.obs
Bingo Machine:

    option explicit: randomize:

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

[Windows] x64 Native Tools Command Prompt for VS 2019

    > nmake -f makefile.win

[OSX] Xcode gcc

    $ make -f makefile.osx

[Linux] GNU gcc

    $ make -f makefile.linux

## Usage
    $ ./oscript sample.obs
    $ echo "wscript.echo 123" | ./oscript

[more...](https://github.com/yrm006/openvbs/blob/master/readme.txt)

## Licence
[CC-BY](http://creativecommons.org/licenses/by/4.0/)

## Author
[NaturalStyle PREMIUM](https://p.na-s.jp)

---
yrm.20191209
