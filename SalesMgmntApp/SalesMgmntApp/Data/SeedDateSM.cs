using Microsoft.EntityFrameworkCore;
using SalesMgmntApp.Entities;
using System;
using System.Reflection.Emit;
using System.Text.RegularExpressions;

namespace SalesManagementApp.Data
{
    public static class SeedData
    {

        public static void AddBloodLineData(ModelBuilder modelBuilder)
        {
            //Add Bird BloodLines
            modelBuilder.Entity<BloodLine>().HasData(new BloodLine
            {

                BloodLineId = 0,
                BloodLineName = "Not Available",
                Status = 0
            });

            modelBuilder.Entity<BloodLine>().HasData(new BloodLine
            {
                BloodLineId = 1,
                BloodLineName = "Not Available",
                Status = 0
            });

            modelBuilder.Entity<BloodLine>().HasData(new BloodLine
            {
                BloodLineId = 2,
                BloodLineName = "Malone",
                Status = 0
            });

            modelBuilder.Entity<BloodLine>().HasData(new BloodLine
            {
                BloodLineId = 3,
                BloodLineName = "Jannsen",
                Status = 0
            });
        }


        //--------------------------------------------------------------------------------







        //--------------------------------------------------------------------------------

        public static void AddBirdData(ModelBuilder modelBuilder)
        {
            //Add Birds
            //With ColourId & BloodLineID
            modelBuilder.Entity<Bird>().HasData(new Bird
            {
                BirdId = 1,
                RingYr = 2014,
                RingMark = "NRC",
                RingNo = "3322",
                BirdName = "Secret Weapon",
                DOB = DateTime.Parse("01 Jun 2014"),
                Gender = "Hen",
                Breeder = "Kieran Malone",
                BredFor = "Kieran Malone",
                BloodLineId = 2,
                ColourId = 1,
                Eye_Code = "NA",
                IsSOH = true,
                IsOLR = false,
                IsBP = false,
                ImagePath = "/Images/Profile/GenericPigeon.jpg",
                SireRingYr = 2014,
                SireRingMark = "NRC",
                SireRingNo = "3322",
                DamRingYr = 2014,
                DamRingMark = "NRC",
                DamRingNo = "3322"

            });
        }
         
        public static void AddLoftData(ModelBuilder modelbuilder)
        {
            //Add Loft
            modelbuilder.Entity<Loft>().HasData(new Loft
                {
                LoftId = 0,
                LoftName = "Not Available",
                FirstName = "Joe",
                LastName = "Bloggs",
                Phone1 = "000000000",
                eMAIL = "joe.bloggs@company.com",
                LicCode="xxx",
                SerialNo="Demo123",
                LastBackupDT = DateTime.Parse("01 Jan 1900"),
                LastRunDT = DateTime.Parse("01 Jan 1900"),
            });
        }




    }

}
