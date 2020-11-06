# Philosophy

SpuriousEmu aims at *easing* the process of analyzing Office documents that are potentially malicious, answering the question

> *Is this file malicious, and if so, what is it doing exactly ?*

However this is not an easy question, as, *in principle*, addressing it would require analyzing all of the possible execution paths of the document, running them in all the possible simulated environments that it can be found in.

As simulating the whole Office ecosystem would not be really efficient, SpuriousEmu is developed with the following philosophy in mind

> *Don't try to achieve the impossible.*

In other words, SpuriousEmu won't answer the question it is made to answer, but it will allow malware analysts to answer it faster. This translates to the following design principles

 1. Features should be robust instead of trying to be too fancy
 2. The output should be integrable with other tools, while staying understandable to a human
 3. Use as little dynamic analysis as possible
 4. When dynamic analysis is unavoidable, don't over-simulate the runtime environment of the document, and let the user do it if possible