vbExSc1.dll is only needed when opening the project from "vbExtra (subclassing outside).vbp" (from "vbExtra (subclassing inside).vbp" it is not needed).
If you compile the OCX from "vbExtra (subclassing outside).vbp", vbExSc1.dll will be also needed to be referenced (and used) in the client program.

vbExTyp.tlb is necessary for both "vbExtra (subclassing outside).vbp" and "vbExtra (subclassing inside).vbp" but only when running in souce code. Once the ocx file is compiled, it is not needed to be referenced and used in the client program.