#!/usr/bin/perl
#use warnings;
#prevents Begin block from assigning global
#use strict;

#my $WINDOWS = 0;

use English qw' -no_match_vars ';

use Parse::CSV;
use Array::Compare;
use Getopt::Long;

use List::Util;

use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::Utility;

use Spreadsheet::WriteExcel;

use Math::Round;
use Math::NumberCruncher;
#use Math::FFT;


#use Tk;
use Treatment;


BEGIN {
	$WINDOWS = ($^O eq 'cygwin') || ($^O eq 'MSWin32' );
	
	if($WINDOWS) {
 		eval 'use Win32::GUI()';
	}
};

my $file_name;
my $artesia_style = '';
my $time_left_stages ='';
my $gui ='';

GetOptions('artesia' => \$artesia_style,'file_name=s' => \$file_name,'timed_volume=s' => \$time_left_stages,'win' => \$gui);


if(1) {
# Win32::MsgBox("HELLO");
if($WINDOWS) {

 my $winMain = Win32::GUI::Window->new(
    -name => 'winMain',
    -text => 'Extracted Volumes',
    -size => [320,240],
); 

my $button = $winMain->AddButton(
    -name => 'Accept',
    -text => 'Accept',
    -pos => [200,120],
);
# These radio buttons are in one group
my $bfield_radio = $winMain->AddRadioButton(
    -name => 'bfield_rad',
    -text => 'Brownfield Style',
    -pos => [10,10],
    -group => 1,
);
my $cog_radio = $winMain->AddRadioButton(
    -name => 'cog_rad',
    -text => 'COG Style',
    -pos => [10,30],
);

$bfield_radio->Checked(1);

sub Accept_Click {

   $winMain->Hide();

   if($bfield_radio->GetCheck()) {
     $artesia_style = undef;	
   }

   if($cog_radio->GetCheck()) {
     $artesia_style = 1;	
   }

   return -1;
}

sub winMain_Terminate { exit; }

$winMain->Show();
Win32::GUI::Dialog();
#http://www.perlmonks.org/?node_id=884217


 $file_name = Win32::GUI::GetOpenFileName(
   -filter =>
     [ 'CSV - Comma Separated Values', '*.csv',
     ],
    -title => 'Select IFS CSV file',
    -filemustexist => 1,
 );
}
else {
print "Current OS does not support --gui option!\n";
exit;
}
}

  my $adi_csv = Parse::CSV->new(
      file => $file_name
  );

  my @valid_data_id = ('Time','Treatment At Wellhead','Stage At Wellhead','Treating Pressure','Slurry Rate','Stage Slurry Vol','Stage Clean Vol','Stage Proppant','Treatment Clean Volume','BH Proppant Conc','Slurry Proppant Conc');
  my $cmp = Array::Compare->new;

  my $i = 0;
  my $valid_data = 0;

  my ($trt_i, $stg_i, $tpres_i, $rate_i, $sv_i, $cv_i, $sp_i,$tcv_i,$bpc_i,$spc_i) = (1..10);

  my %job_treatments;

  my $trigger = 0;

#  my $min_ball_stage = 1;
#  my $max_ball_stage = 4;
#  my $min_ball_rate = 10;
#  my $max_ball_rate = 15;
#  my $ball_at_rate = 30;
#  my $start_ball_count = 0;
#  my $is_ball_with_stage = 0;
#  my $is_rate_dropped_ball = 0;
#  my $within_rate_ball = 0;


  my $min_ball_stage = 1;
  my $max_ball_stage = 4;

  my $ball_pre_rate = 0;
  my $ball_at_rate = 30;

  my $landing_rate = 16;

  my $is_ball_with_stage = 0;
  my $is_rate_dropped = 0;
  my $ball_calc_end = 0;


  my $curr_treatment_clean_vol = 0;

  while ( my $line_arr_ref = $adi_csv->fetch ) {
	  if($cmp->compare( \@valid_data_id, $line_arr_ref) and not $valid_data) {
	    #burn a line
	    $adi_csv->fetch;
	    $valid_data = 1;
	  }
	  else {
	    if($valid_data) {

		$curr_treatment_clean_vol = $$line_arr_ref[$tcv_i];

		if ( not (exists $job_treatments{ $$line_arr_ref[$trt_i] } ) ) {
		#if treatment does not exists, add stage, cv, prop
		   my $curr_treatment = \Treatment->new();
		   $job_treatments{$$line_arr_ref[$trt_i]} = $curr_treatment; 

		   bless($curr_treatment,'Treatment');

		   my $curr_stage_vol = $$curr_treatment -> stage_clean_vol();
		   my $curr_stage_prop = $$curr_treatment -> stage_prop_vol();

		   my $curr_stage_svol = $$curr_treatment -> stage_slurry_vol();
		   my $curr_stage_press = $$curr_treatment -> stage_pressures();
		   my $curr_stage_rates = $$curr_treatment -> stage_rates();

		   my $curr_treatment_pressure = $$curr_treatment -> treatment_pressure();
		   my $curr_treatment_rate = $$curr_treatment -> treatment_rate();

		   my $curr_treatment_max_pc = $$curr_treatment -> treatment_max_pc();

		   my $curr_ball_press = $$curr_treatment -> ball_calc_pressure();

		   my $curr_ball_rate = $$curr_treatment -> ball_calc_rate();

 		   $$curr_stage_vol{$$line_arr_ref[$stg_i]} = $$line_arr_ref[$cv_i]; 
		   $$curr_stage_prop{$$line_arr_ref[$stg_i]} = $$line_arr_ref[$sp_i]; 

		   $$curr_stage_svol{$$line_arr_ref[$stg_i]} = $$line_arr_ref[$sv_i];

		   $$curr_stage_press{$$line_arr_ref[$stg_i]} = [];

		   $curr_treatment_pressure = [];
		   $curr_treatment_rate = [];

		   $curr_ball_press = [];
		   $curr_ball_rate = [];

		   my $press_arr_ref = $$curr_stage_press{$$line_arr_ref[$stg_i]}; 
		   push @{$press_arr_ref}, $$line_arr_ref[$tpres_i]; 

		   $$curr_stage_rates{$$line_arr_ref[$stg_i]} = [];

		   my $rates_arr_ref = $$curr_stage_rates{$$line_arr_ref[$stg_i]}; 
		   push @{$rates_arr_ref}, $$line_arr_ref[$rate_i]; 

		   $$curr_treatment_max_pc = 0;

		   $trigger = 0;
  		   $is_rate_dropped = 0;
		   $ball_calc_end = 0;
                   $is_ball_with_stage = 0;

		}
		else {
		#if treatment exists, check if stage exists
		   
		   my $curr_treatment = $job_treatments{$$line_arr_ref[$trt_i]};
		   
		   bless($curr_treatment,'Treatment');

		   my $curr_stage_vol = $$curr_treatment -> stage_clean_vol();
		   my $curr_stage_prop = $$curr_treatment -> stage_prop_vol();

		   my $curr_stage_svol = $$curr_treatment -> stage_slurry_vol();

		   my $curr_stage_press = $$curr_treatment -> stage_pressures();
		   my $curr_stage_rates = $$curr_treatment -> stage_rates();


		   my $curr_treatment_pressure = $$curr_treatment -> treatment_pressure();
		   my $curr_treatment_rate = $$curr_treatment -> treatment_rate();

		   my $curr_treatment_max_pc = $$curr_treatment -> treatment_max_pc();

		   my $curr_ball_press = $$curr_treatment -> ball_calc_pressure();
		   my $curr_ball_rate = $$curr_treatment -> ball_calc_rate();

		   if ( $$line_arr_ref[$bpc_i] >= 0.25  and not $trigger)  {
		     $trigger = 1;
				   $$curr_treatment_max_pc = $$line_arr_ref[$spc_i];
				   #print $$curr_treatment_max_pc," --\n";

	           }

		   else {
			   if ( $$line_arr_ref[$bpc_i] < 0.10  and $trigger)  {
				$trigger = 0;
			   }
			}

		   if ($trigger) {
			   push @{$curr_treatment_pressure}, $$line_arr_ref[$tpres_i];
			   push @{$curr_treatment_rate}, $$line_arr_ref[$rate_i];

			   if($$line_arr_ref[$spc_i] > $$curr_treatment_max_pc ) {
				   $$curr_treatment_max_pc = $$line_arr_ref[$spc_i];
				   #print $$curr_treatment_max_pc,"\n";
			   }
		   }

		   if ($$line_arr_ref[$stg_i] >= $min_ball_stage && $$line_arr_ref[$stg_i] <= $max_ball_stage  && not $is_ball_with_stage) {
			   $is_ball_with_stage = 1;
		   }

		   if ($is_ball_with_stage and $$line_arr_ref[$rate_i] >= $ball_at_rate and not $ball_pre_rate) {
			   $ball_pre_rate = 1;
		   }

		   if ($ball_pre_rate and $$line_arr_ref[$rate_i] < $ball_at_rate and not $is_rate_dropped) {
			   if ($$line_arr_ref[$rate_i] <= $landing_rate) {
				   $is_rate_dropped = 1;
			   }
		   }

		   if ($is_rate_dropped and $$line_arr_ref[$rate_i] > $landing_rate and not $ball_calc_end) {
			   $ball_calc_end = 1;
		   }

		   if (not $ball_calc_end and $is_rate_dropped) {
			   push @{$curr_ball_press}, $$line_arr_ref[$tpres_i];
		   }


		   if ( not (exists $$curr_stage_vol{ $$line_arr_ref[$stg_i] } ) )  {
		     #if stage does not exist, add cv and prop
 		   
			  $$curr_stage_vol{$$line_arr_ref[$stg_i]} = $$line_arr_ref[$cv_i]; 
		          $$curr_stage_prop{$$line_arr_ref[$stg_i]} = $$line_arr_ref[$sp_i]; 

		          $$curr_stage_svol{$$line_arr_ref[$stg_i]} = $$line_arr_ref[$sv_i];

		          $$curr_stage_press{$$line_arr_ref[$stg_i]} = [];

		          my $arr_ref = $$curr_stage_press{$$line_arr_ref[$stg_i]}; 
		          push @{$arr_ref}, $$line_arr_ref[$tpres_i]; 


		          $$curr_stage_rates{$$line_arr_ref[$stg_i]} = [];

		          my $rates_arr_ref = $$curr_stage_rates{$$line_arr_ref[$stg_i]}; 
		          push @{$rates_arr_ref}, $$line_arr_ref[$rate_i]; 


		   }
		   else {
		     #stage exists check if vols are higher then increment
		     if( $$curr_stage_vol{$$line_arr_ref[$stg_i]} < $$line_arr_ref[$cv_i] ) {
			     $$curr_stage_vol{$$line_arr_ref[$stg_i]} = $$line_arr_ref[$cv_i];
		     }

		     if( $$curr_stage_prop{$$line_arr_ref[$stg_i]} < $$line_arr_ref[$sp_i] ) {
			     $$curr_stage_prop{$$line_arr_ref[$stg_i]} = $$line_arr_ref[$sp_i];
		     }

		     if( $$curr_stage_svol{$$line_arr_ref[$stg_i]} < $$line_arr_ref[$sv_i] ) {
			     $$curr_stage_svol{$$line_arr_ref[$stg_i]} = $$line_arr_ref[$sv_i];
		     }

		     my $arr_ref = $$curr_stage_press{$$line_arr_ref[$stg_i]}; 
		     push @{$arr_ref}, $$line_arr_ref[$tpres_i]; 

		     my $rates_arr_ref = $$curr_stage_rates{$$line_arr_ref[$stg_i]}; 
		     push @{$rates_arr_ref}, $$line_arr_ref[$rate_i]; 

		   }
		}
             }
	  }
  }


#find absolute max stage number from all treatments
my $abs_max_stgs = 0;
foreach my $t_i (sort keys %job_treatments) {
 my $tmnt = {}; 

 bless($tmnt,qw(Treatment));

 $tmnt = $job_treatments{$t_i};

 my %stg_vol = %{$$tmnt->stage_clean_vol()};
 my @stgs = keys %stg_vol;

 foreach my $s (@stgs) {
   if($abs_max_stgs < $s) {
	$abs_max_stgs = $s;
   }
   
 }

}


if ( $time_left_stages ) {

	my ($stage_file,$t_num) = split(/,/,$time_left_stages);
	$t_num = 0 unless defined $t_num;

	my $prsr = Spreadsheet::ParseExcel->new();
	my $designed_stage_vols = $prsr->parse($stage_file);
	my $wsht = $designed_stage_vols->worksheet("stages");

	#need to check if valid worksheet and parser and check for a default sheet;

	my @designed_treatment;
        
	my ( $row_min, $row_max ) = $wsht->row_range();

	my $cum_vol = 0;

	for my $i ($row_min ... $row_max) {
		my $cell = $wsht->get_cell($i,0);
		$cum_vol += $cell->value();
		push @designed_treatment, $cell->value();
	}

        my $tmnt = {}; 

        bless($tmnt,qw(Treatment));

        $tmnt = $job_treatments{$t_num};
        my %stage_vol = %{$$tmnt -> stage_clean_vol()};

        my @stgs_sorted = keys %stage_vol;
        @stgs_sorted = sort { $a <=> $b } @stgs_sorted;

	my @stages_to_add = @stgs_sorted;

	pop @stages_to_add;

	foreach my $s (@stages_to_add) {
		$designed_treatment[$s-1] = $stage_vol{$s};
	}

	$cum_vol = 0;
	foreach my $s (@designed_treatment) {
		$cum_vol +=$s;
	}
	print round($cum_vol),"\n";


	exit;
}

if ( not $artesia_style) {
# Create a new Excel workbook
my $new_wbook = Spreadsheet::WriteExcel->new('extracted_volumes.xls');

die "Problems creating new Excel file: $!" unless defined $new_wbook;

# Add a worksheet
my $write_wsheet = $new_wbook->add_worksheet('extracted_volumes');
my $new_row = 0;
my $new_col = 1;
my $add_stage = 1;

for (1..$abs_max_stgs) {
 $write_wsheet->write(1+$_, 0, $add_stage++);
}

my @treatments_sorted = keys %job_treatments;
@treatments_sorted = sort { $a <=> $b } @treatments_sorted;

$write_wsheet->write($abs_max_stgs + 5, 0, "tmnt #");
$write_wsheet->write($abs_max_stgs + 5, 1, "avg p");
$write_wsheet->write($abs_max_stgs + 5, 2, "avg r");
$write_wsheet->write($abs_max_stgs + 5, 3, "max p");
$write_wsheet->write($abs_max_stgs + 5, 4, "max r");
$write_wsheet->write($abs_max_stgs + 5, 5, "max pc");

foreach my $t_i (@treatments_sorted) {
 my $tmnt = {}; 

 bless($tmnt,qw(Treatment));

 $tmnt = $job_treatments{$t_i};

 my $curr_tp = $$tmnt -> treatment_pressure();
 my $curr_rate = $$tmnt -> treatment_rate();
 my $curr_max_pc = $$tmnt -> treatment_max_pc();

 my @curr_ball_tp = @{$$tmnt -> ball_calc_pressure()}; ##needs to be a power of 2 make a check


 if(scalar(@curr_ball_tp) > 0) {
# push @curr_ball_tp, (0) x (4096-scalar(@curr_ball_tp));
#print "STD: ", Math::NumberCruncher::StandardDeviation(\{@curr_ball_tp[(1000..1500)]}),"\n";

#my @tmp =  @curr_ball_tp[(1500..2000)];
#print "STD: ", Math::NumberCruncher::StandardDeviation(\@tmp),"\n";

#foreach my $b (@curr_ball_tp) {
# foreach my $b (@tmp) {
#		print "$b\n";
# }
#

	 for(my $i = 0; $i<scalar(@curr_ball_tp); $i+=20) {
		 #for ( my $j = $i; $j < $i+20 and $j< scalar(@curr_ball_tp); $j++) {
		 #  print $j," ",$curr_ball_tp[$j],"\n";
		 #}
		 #print "\n";
		 #

		 
	 }

 }


 ###test FFT###
# my $curr_ball_tp = $$tmnt -> ball_calc_pressure(); ##needs to be a power of 2 make a check
 
# push @$curr_ball_tp, (0) x (4096-scalar(@$curr_ball_tp));
# my $fft = new Math::FFT(\@curr_ball_tp);
 
 #my $spectrum = $fft->spctrm;

# foreach my $x (@$fft) {
#	print "$x\n";
 #}

# my $spectrum_peak = List::Util::max(@$spectrum);

# my $idx = List::Util::first { $$spectrum[$_] ge $spectrum_peak } 0..scalar(@$spectrum);

# print "$idx and $spectrum_peak\n";


#averages/max
 if( scalar @$curr_tp > 0) {
	$write_wsheet->write($abs_max_stgs + 5 + $new_col, $new_row, $t_i);
	$write_wsheet->write($abs_max_stgs + 5 + $new_col, $new_row + 1,  round(Math::NumberCruncher::Mean($curr_tp)));
	$write_wsheet->write($abs_max_stgs + 5 + $new_col, $new_row + 2,  sprintf("%.2f",Math::NumberCruncher::Mean($curr_rate)));
	$write_wsheet->write($abs_max_stgs + 5 + $new_col, $new_row + 3,  (round(Math::NumberCruncher::Range($curr_tp)))[0]);
	$write_wsheet->write($abs_max_stgs + 5 + $new_col, $new_row + 4,  sprintf("%.2f",(Math::NumberCruncher::Range($curr_rate))[0]));
	$write_wsheet->write($abs_max_stgs + 5 + $new_col, $new_row + 5,  sprintf("%.2f",$$curr_max_pc));
}

 $write_wsheet->write($new_row,  $new_col ,"Tmnt #$t_i");
 $new_row++;

 $write_wsheet->write($new_row,  $new_col ,"SC (gal)");
 $write_wsheet->write($new_row,  $new_col+1 ,"SP (lb)");
 $new_row++;

 my %stage_vol = %{$$tmnt -> stage_clean_vol()};
 my %stage_prop = %{$$tmnt -> stage_prop_vol()};

 my @stgs_sorted = keys %stage_vol;
 @stgs_sorted = sort { $a <=> $b } @stgs_sorted;

 my %skipped_cv_stages;
 my %skipped_sp_stages;

 foreach my $s_i (@stgs_sorted) {
   if($s_i eq $new_row-1) {
     $write_wsheet->write($new_row, $new_col, round($stage_vol{$s_i}));
     $write_wsheet->write($new_row, $new_col+1, round($stage_prop{$s_i}));
   }
   else {
     $skipped_cv_stages{$s_i} = $stage_vol{$s_i};
     $skipped_sp_stages{$s_i} = $stage_prop{$s_i};
   }
   $new_row++;
 }

 my @skipped_stgs = keys %skipped_cv_stages;
 @skipped_stgs = sort { $a <=> $b } @skipped_stgs;
 my $total_skipped = scalar(@skipped_stgs);

 if($total_skipped > 0) {
   my $sk = 1;
   for($new_row = 2; $new_row - 1 <= $abs_max_stgs; $new_row++) {
      if($sk eq $new_row-1) {
        $write_wsheet->write($new_row, $new_col, round($stage_vol{$sk}));
        $write_wsheet->write($new_row, $new_col+1, round($stage_prop{$sk}));
      }
      $sk++;
   }
 }

 $new_row=0;
 $new_col+=2;

}
$new_wbook->close();
Win32::MsgBox("Finished!",0,"Check your folder.");
}

else {
  my @header = ('Avg Pressure','Stage Slurry','Slurry BBL','Avg Rate','Stage Clean','Clean BBL','Mass Prop');
  my @units = ('psi','gal','bbl','bpm','gal','bbl','lb');
  # Create a new Excel workbook
  my $new_wbook = Spreadsheet::WriteExcel->new('extracted_volumes.xls');
  die "Problems creating new Excel file: $!" unless defined $new_wbook;

  my @treatments_sorted = keys %job_treatments;
  
  @treatments_sorted = sort { $a <=> $b } @treatments_sorted;

  foreach my $t_i (@treatments_sorted) {
    # Add a worksheet
    my $write_wsheet = $new_wbook->add_worksheet("Treatment #$t_i");
    my $new_row = 0;
    my $new_col = 1;
    my $add_stage = 1;
    my $j =1;

    foreach (@header) {
      $write_wsheet->write($new_row, $j++, $_);
    }
    $j=1;
    $new_row++;
    foreach (@units) {
      $write_wsheet->write($new_row, $j++, $_);
    }
   
    $new_row++;
    my $tmnt = {}; 

    bless($tmnt,qw(Treatment));

    $tmnt = $job_treatments{$t_i};

    my $tmnt_max_stgs = 0;

    my %stage_vol = %{$$tmnt -> stage_clean_vol()};
    my %stage_prop = %{$$tmnt -> stage_prop_vol()};

    my %stage_svol = %{$$tmnt -> stage_slurry_vol()};

    my %stage_press = %{$$tmnt -> stage_pressures()};
    my %stage_rates = %{$$tmnt -> stage_rates()};

    my @stgs_sorted = keys %stage_vol;
    @stgs_sorted = sort { $a <=> $b } @stgs_sorted;

    $tmnt_max_stgs = $stgs_sorted[-1];
 
    my %skipped_cv_stages;
    my %skipped_sp_stages;
    my %skipped_sv_stages;
    my %skipped_tp_stages;

    for (1..$tmnt_max_stgs) {
      $write_wsheet->write(1+$_, 0, $add_stage++);
    }


    $write_wsheet->write($abs_max_stgs + 10, 0, "tmnt #");
    $write_wsheet->write($abs_max_stgs + 10, 1, "avg p");
    $write_wsheet->write($abs_max_stgs + 10, 2, "avg r");
    $write_wsheet->write($abs_max_stgs + 10, 3, "max p");
    $write_wsheet->write($abs_max_stgs + 10, 4, "max r");
    $write_wsheet->write($abs_max_stgs + 10, 5, "max pc");
    
    my $curr_tp = $$tmnt->treatment_pressure();
    my $curr_rate = $$tmnt->treatment_rate();
    my $curr_max_pc = $$tmnt->treatment_max_pc();

#averages/max
 if( scalar @$curr_tp > 0) {
	$write_wsheet->write($abs_max_stgs + 11, 0, $t_i);
	$write_wsheet->write($abs_max_stgs + 11, 1, round(Math::NumberCruncher::Mean($curr_tp)));
	$write_wsheet->write($abs_max_stgs + 11, 2, sprintf("%.2f",Math::NumberCruncher::Mean($curr_rate)));
	$write_wsheet->write($abs_max_stgs + 11, 3, (round(Math::NumberCruncher::Range($curr_tp)))[0]);
	$write_wsheet->write($abs_max_stgs + 11, 4, sprintf("%.2f",(Math::NumberCruncher::Range($curr_rate))[0]));
	$write_wsheet->write($abs_max_stgs + 11, 5, sprintf("%.2f",$$curr_max_pc));
}

    foreach my $s_i (@stgs_sorted) {
       if($s_i eq $new_row-1) {
	 no warnings 'uninitialized';
         $write_wsheet->write($new_row, $new_col, round(Math::NumberCruncher::Mean($stage_press{$s_i})));
         $write_wsheet->write($new_row, $new_col+1, round($stage_svol{$s_i}));
         $write_wsheet->write($new_row, $new_col+2, round($stage_svol{$s_i}/42.0));
         $write_wsheet->write($new_row, $new_col+3, sprintf("%.2f",Math::NumberCruncher::Mean($stage_rates{$s_i})));
         $write_wsheet->write($new_row, $new_col+4, round($stage_vol{$s_i}));
         $write_wsheet->write($new_row, $new_col+5, round($stage_vol{$s_i}/42.0));
         $write_wsheet->write($new_row, $new_col+6, round($stage_prop{$s_i}));
       }
       else {
         $skipped_cv_stages{$s_i} = $stage_vol{$s_i};
         $skipped_sp_stages{$s_i} = $stage_prop{$s_i};
         $skipped_sv_stages{$s_i} = $stage_svol{$s_i};
         $skipped_tp_stages{$s_i} = Math::NumberCruncher::Mean($stage_press{$s_i});
         $skipped_tp_stages{$s_i} = Math::NumberCruncher::Mean($stage_rates{$s_i});
       }
       $new_row++;
    }

    my @skipped_stgs = keys %skipped_cv_stages;
    @skipped_stgs = sort { $a <=> $b } @skipped_stgs;
    my $total_skipped = scalar(@skipped_stgs);

    if($total_skipped > 0) {
      my $sk = 1;
      for($new_row = 2; $new_row -1 <= $abs_max_stgs; $new_row++) {
         if($sk eq $new_row-1) {
	   no warnings 'uninitialized';
           $write_wsheet->write($new_row, $new_col, round(Math::NumberCruncher::Mean($stage_press{$sk})));
           $write_wsheet->write($new_row, $new_col+1, round($stage_svol{$sk}));
           $write_wsheet->write($new_row, $new_col+2, round($stage_svol{$sk}/42.0));
           $write_wsheet->write($new_row, $new_col+3, sprintf("%.2f",Math::NumberCruncher::Mean($stage_rates{$sk})));
           $write_wsheet->write($new_row, $new_col+4, round($stage_vol{$sk}));
           $write_wsheet->write($new_row, $new_col+5, round($stage_vol{$sk}/42.0));
           $write_wsheet->write($new_row, $new_col+6, round($stage_prop{$sk}));
         }
        $sk++;
      }
    }

    $new_row=1;
    $new_col+=7;

   }
$new_wbook->close();
Win32::MsgBox("Finished!",0,"Check your folder.");
 }
