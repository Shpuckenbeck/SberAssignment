using System;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Infrastructure;
using Microsoft.EntityFrameworkCore.Metadata;
using Microsoft.EntityFrameworkCore.Migrations;
using Sberbank.Data;

namespace Sberbank.Migrations
{
    [DbContext(typeof(BankContext))]
    [Migration("20180215131517_initial")]
    partial class initial
    {
        protected override void BuildTargetModel(ModelBuilder modelBuilder)
        {
            modelBuilder
                .HasAnnotation("ProductVersion", "1.1.5")
                .HasAnnotation("SqlServer:ValueGenerationStrategy", SqlServerValueGenerationStrategy.IdentityColumn);

            modelBuilder.Entity("Sberbank.Models.Record", b =>
                {
                    b.Property<int>("RecordId")
                        .ValueGeneratedOnAdd();

                    b.Property<float>("currency");

                    b.Property<DateTime>("date");

                    b.Property<double>("earnings");

                    b.Property<float>("index");

                    b.HasKey("RecordId");

                    b.ToTable("Records");
                });
        }
    }
}
