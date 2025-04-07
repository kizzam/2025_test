using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    public partial class Race
    {
        public DateTime? Reldate { get; set; }
        public string? CodePnt { get; set; }
        public string? CodeRace { get; set; }
        public int? Relhrs { get; set; }
        public int? Relmins { get; set; }
        public bool? Indampm { get; set; }
        public DateTime? ClkDate { get; set; }
        public int? ClkHrs { get; set; }
        public int? ClkMins { get; set; }
        public int? ClkSecs { get; set; }
        public float? FlyTime { get; set; }
        public DateTime? Bskdate { get; set; }
        public string? Reldatet { get; set; }
        public string? Position { get; set; }
        public string? RaceNotes { get; set; }
        public string? RaceWeather { get; set; }
        public string? Status { get; set; }
        public int? ClubPos { get; set; }
        public float? ClubVel { get; set; }
        public int? ClubPoints { get; set; }
        public double? ClubMinsDarkness { get; set; }
        public float? ClubFtime { get; set; }
        public int? FedPos { get; set; }
        public float? FedVel { get; set; }
        public int? FedPoints { get; set; }
        public int? FedMinsDarkness { get; set; }
        public float? FedFtime { get; set; }
        public int? NatPos { get; set; }
        public int? IntNatPos { get; set; }
        public int? Yr { get; set; }
        public string? Mark { get; set; }
        public string? Ringno { get; set; }
        public string? Spare { get; set; }
        public string? Spare1 { get; set; }
        public int? Code { get; set; }
        public double? Spare2 { get; set; }
        public int RaceCode { get; set; }
        public int? Ooclub { get; set; }
        public int? Oosec { get; set; }
        public int? Oofed { get; set; }
        public int? Oocomb { get; set; }
        public int? Oonat { get; set; }
        public int? Ooinat { get; set; }
        public int? BiClub { get; set; }
        public int? BiComb { get; set; }
        public int? BiFed { get; set; }
        public int? BiSec { get; set; }
        public int? BiNat { get; set; }
        public int? BiInat { get; set; }
        public int? BiLoft { get; set; }
        public string? RaceStatus { get; set; }
        public int? Ooloft { get; set; }
        public float? WinVel { get; set; }
        public double? Distance { get; set; }
    }
}
