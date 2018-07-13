###############################################################################
## OCSINVENTORY-NG
## Copyleft Stephane PAUTREL 2018
## Web : http://www.acb78.com
##
## This code is open source and may be copied and modified as long as the source
## code is always made freely available.
## Please refer to the General Public Licence http://www.gnu.org/ or Licence.txt
################################################################################
 
package Apache::Ocsinventory::Plugins::Driverslist::Map;
 
use strict;
 
use Apache::Ocsinventory::Map;

$DATA_MAP{driverslist} = {
	mask => 0,
		multi => 1,
		auto => 1,
		delOnReplace => 1,
		sortBy => 'CLASS',
		writeDiff => 0,
		cache => 0,
		fields => {
			CLASS => {},
			DESCRIPTION => {},
			NAME => {},
			PROVIDERNAME => {},
			VERSION => {},
			DATE => {},
			INFNAME => {},
			FRIENDLYNAME => {},
			SIGNED => {},
			DEVICEID => {}
}
};
1;
