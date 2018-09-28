# VBASeismograph
VBA Seismograph is a tool for detecting VBA stomping. It has been
developed and tested under Ubuntu 16.04. This is done by checking for:

* Functions and variables that are defined in the compiled p-code that
  do not appear in the VBA source code.
* String literals that are used in the compiled p-code that do not
  appear in the VBA source code.
* Comments that appear in the compiled p-code that do not appear in
  the VBA source code.

## Installation

VBA Seismograph makes use of two external packages, pcodedmp and
ClamAV's sigtool. To install these (under Ubuntu):

### Install pcodedmp

[pcodemp.py](https://github.com/bontchev/pcodedmp) is a p-code
disassembler. To install it do the following:

```
git clone https://github.com/bontchev/pcodedmp.git
```

### Install ClamAv

ClamAV is an open source AV scanner. It contains a utility called
sigtool that performs (among other things) VBA source code
decompression for Office documents. To install ClamAV under Ubuntu do
the following:

```
sudo apt-get install clamav
```

### PCODEDMP_DIR Environment Variable

VBA Seismograph reads the install directory for pcodedmp from the
PCODEDMP_DIR environment variable. To set this under csh add something
like the following (modified for where you installed pcodedmp) to your
.cshrc file:

```
setenv PCODEDMP_DIR /home/sayre/Software/pcodedmp
```

To set this under bash add something like the following (modified for
where you installed pcodedmp) to your .bashrc file:

```
export PCODEDMP_DIR=/home/sayre/Software/pcodedmp
```

## Usage

To get help run:

```
vba_seismograph.py -h
```

To check the Office file FOO (Excel or Word file) run:

```
vba_seismograph.py FOO
```

To get details about differences between the p-code and the VBA source
code run:

```
vba_seismograph.py -v FOO
```
