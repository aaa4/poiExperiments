package com.myexcel.demo;

import java.time.LocalDate;

public class Unit {


    String unitType;
    String unitNumber;
    LocalDate unitReleaseDate;
    String unitNote;

    public Unit() {
    }

    public Unit(String unitType, String unitNumber, LocalDate unitReleaseDate, String unitNote) {
        this.unitType = unitType;
        this.unitNumber = unitNumber;
        this.unitReleaseDate = unitReleaseDate;
        this.unitNote = unitNote;
    }


    public String getUnitType() {
        return unitType;
    }

    public void setUnitType(String unitType) {
        this.unitType = unitType;
    }

    public String getUnitNumber() {
        return unitNumber;
    }

    public void setUnitNumber(String unitNumber) {
        this.unitNumber = unitNumber;
    }

    public LocalDate getUnitReleaseDate() {
        return unitReleaseDate;
    }

    public void setUnitReleaseDate(LocalDate unitReleaseDate) {
        this.unitReleaseDate = unitReleaseDate;
    }

    public String getUnitNote() {
        return unitNote;
    }

    public void setUnitNote(String unitNote) {
        this.unitNote = unitNote;
    }


    @Override
    public String toString() {
        return "Unit{" +
                "unitType='" + unitType + '\'' +
                ", unitNumber='" + unitNumber + '\'' +
                ", unitReleaseDate=" + unitReleaseDate +
                ", unitNote='" + unitNote + '\'' +
                '}';
    }
}
