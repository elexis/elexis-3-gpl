## Overview

In this repository you find features for Elexis which contain code under the GPL.

We have separted this plugin, as many Elexis user are running the non-EPL spin-off offered by the Medelexis AG.

As the Medelexis AG does (as common sense dictates) want to add any GPL code in their application which contains code,
they are unwilling to distribute under an GPL compatible license, the Elexis developer search for a way to allow
Medelexis clients to use these plugins. Medelexis version 3.1.0  or higher allow placing additional features under
in a dropins subdirectory. Therefore (Med-)Elexis users just download the feature and plugins folder from a branch
under https://download.elexis.info/elexis.3.gpl/, place them in their (Med-)Elexis application folder into a
(newly created) subdirectory called dropins.

### Build from source


You will need java8 and maven >= 3.3

     git clone https://github.com/elexis/elexis-3-gpl
     cd elexis-3-gpl
     mvn -V clean verify -Dtycho.localArtifacts=ignore -Pall-archs
