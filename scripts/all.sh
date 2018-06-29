#/bin/sh

set -e

scripts/load_companies.pl
scripts/customers.pl
scripts/programs.pl
scripts/orders.pl
scripts/order_detail.pl
scripts/order_assoc.pl
