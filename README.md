# KARTS
KARTS (Kate's Automated Ready-To-Ship) performs a series of transformations on EDI order data, from checking prices against a list and calculating ship dates to validating addresses in the FedEx API and sending emails. It also generates SQL statements to modify order data that exists outside of the Order_Import table, which is needed to fully complete an order record in the ERP and database combination we're working with.

KARTS has several subcomponents:

WebMapper:
Converts Squarespace orders to the mapping the main script is expecting, and has some logic to deal with commonly encountered issues (such as indexes shifting, and blank lines in the export file.)

FreightQuoter: 
Includes a more robust version of the packing method howManyFit, and can convert user-supplied data into a FedEx shipping quote.

810Checker:
Validates outgoing EDI 810 invoice documents before they are sent to the customer.

KARTS was created for TrippNT Carts in collaboration with Sam Skiver, Kyle Malios, and Jason Ayers.
