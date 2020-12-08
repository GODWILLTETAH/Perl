 #Perl excel writer module
use Excel::Writer::XLSX;

my $workbook = Excel::Writer::XLSX->new ('file1.xlsx'); #creates a new workbook
$workbook->set_properties( #set properties
    title    => 'IncludeHelp',
    author   => 'Grace',
    comments => 'first edition',
);
my $worksheet = $workbook->add_worksheet ('page1'); #add worksheets
my $worksheet = $workbook->add_worksheet ('page2');
my $worksheet = $workbook->add_worksheet ('page3');

for $worksheet ( $workbook->sheets() ) { #loop through all the sheets and write
    $worksheet->write( 'A1', 'Hello' );
    print $worksheet->get_name();  #get and print the name of each worksheet
    print "\n";
}
$workbook->close(); #close workbook
print 'done!';