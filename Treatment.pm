package Treatment;

use warnings;
use strict;
use Carp;

my $this_stage_clean_vol;
my $this_stage_prop_vol;

my $this_stage_slurry_vol;
my $this_stage_press_arr;
my $this_stage_rate_arr;

my $this_treatment_pressure;
my $this_treatment_rate;

my $this_max_prop_con;

my $this_ball_calc_pressure;
my $this_ball_rate;

sub new {
    my $class = shift;

    $this_stage_clean_vol = {};
    $this_stage_prop_vol = {};
    
    $this_stage_slurry_vol = {};
    $this_stage_press_arr = {};
    $this_stage_rate_arr = {}; 

    $this_treatment_pressure = [];
    $this_treatment_rate = [];
    $this_max_prop_con = 0;

    $this_ball_calc_pressure = [];
    $this_ball_rate = [];

    my $self = {
     	stage_vol => $this_stage_clean_vol,
     	stage_prop => $this_stage_prop_vol,

	stage_svol => $this_stage_slurry_vol,
	stage_press => $this_stage_press_arr,
	stage_rate => $this_stage_rate_arr,

	treatment_pressure => $this_treatment_pressure,
	treatment_rate => $this_treatment_rate,
	treatment_max_pc => $this_max_prop_con,
	ball_calc_pressure => $this_ball_calc_pressure,
	ball_calc_rate => $this_ball_rate
    };

    bless($self,$class);
    return $self;
};

sub stage_clean_vol {
	my $self = shift;
	return $self->{stage_vol};
}

sub stage_prop_vol {
	my $self = shift;
	return $self->{stage_prop};
}

sub stage_slurry_vol {
	my $self = shift;
	return $self->{stage_svol};
}

sub stage_pressures {
	my $self = shift;
	return $self->{stage_press};
}

sub stage_rates {
	my $self = shift;
	return $self->{stage_rate};
}

sub treatment_pressure {
	my $self = shift;
	return $self->{treatment_pressure};
}

sub treatment_rate {
	my $self = shift;
	return $self->{treatment_rate};
}

sub treatment_max_pc {
	my $self = shift;
	return \$self->{treatment_max_pc};
}

sub ball_calc_pressure {
	my $self = shift;
	return $self->{ball_calc_pressure};
}

sub ball_calc_rate {
	my $self = shift;
	return $self->{ball_calc_rate};
}

1
