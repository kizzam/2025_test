using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    public partial class Racetran
    {
        public DateTime? Reldate { get; set; }
        public string? CodePnt { get; set; }
        public string? CodeRace { get; set; }
        public int Yr { get; set; }
        public string Mark { get; set; } = null!;
        public string Ringno { get; set; } = null!;
        public double? Velocity { get; set; }
        public int? Position { get; set; }
        public string? Notes { get; set; }
        public DateTime? Clkdate { get; set; }
        public string? Racedatet { get; set; }
        public int? ClkHr { get; set; }
        public int? ClkMin { get; set; }
        public int? ClkSec { get; set; }
        public double? FlyTimeHrs { get; set; }
        public int RaceCode { get; set; }
        public string? Status { get; set; }
        public int? PosClub { get; set; }
        public int? PosLoft { get; set; }
        public int? PosComb { get; set; }
        public int? PosSec { get; set; }
        public int? PosFed { get; set; }
        public int? PosNat { get; set; }
        public int? PosInat { get; set; }
        public int? Ooloft { get; set; }
        public int? Ooclub { get; set; }
        public int? Oocomb { get; set; }
        public int? Oosec { get; set; }
        public int? Oofed { get; set; }
        public int? Oonat { get; set; }
        public int? Ooinat { get; set; }
        public string? Rating { get; set; }
        public string? CodeRating { get; set; }
        public int? Fthr { get; set; }
        public int? Ftmin { get; set; }
        public int? Ftsec { get; set; }
        public string? RaceStatus { get; set; }
        public double? Distance { get; set; }
    }
}
