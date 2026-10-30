# Appointment & Booking Management System (Excel + VBA)

An Excel-based appointment and booking management system built with VBA, PivotTables and an interactive dashboard. The example data is for a **salon**, but the same structure can be adapted for other businesses such as an institute or a clinic by changing the services, staff and data.

> **Note:** All data in this workbook is dummy data created for demonstration.

---

## Features

### Booking Form (VBA UserForm)
- **Auto Appointment ID**: a new ID (for example `A0501`) is generated for every booking
- **Duplicate time slot validation**: warns the user if the same date and time is already booked
- **Customer search**: select a customer to load their booking details into the form
- **Auto status**: every new booking is saved with the status "Booked"
- **PDF appointment slip**: creates a PDF slip for the booking shown in the form
- **Print booking receipt**: opens a print preview of the booking receipt

### Dashboard
- KPI cards: Total Appointments, Today's Appointments, Completed, Cancelled, Monthly Revenue, Pending Payments
- Charts: Status (donut), Staff-wise Appointments, Service-wise Appointments
- **Slicers** for Service Name, Status and Staff Name to filter the charts interactively
- **Refresh Dashboard** button to update everything after new bookings

### Other
- Status colour coding using Conditional Formatting (no VBA)
- Master sheets and reports built on the booking data

---

## Tools and Techniques Used
- Microsoft Excel (macro-enabled workbook, `.xlsm`)
- VBA (UserForms, events, validation, PDF export, print preview)
- PivotTables, Pivot Charts and Slicers
- Conditional Formatting
- Excel formulas for KPIs

---

## How to Use

1. Download `Appointment_Booking_System12.xlsm` from this repository.
2. Open it in **desktop Microsoft Excel** (macros do not run in Excel Online, Google Sheets or phone apps).
3. Click **Enable Content** in the yellow security bar at the top.
4. Open the booking form from the menu / button in the workbook and try:
   - saving a new booking,
   - booking the same slot again to see the duplicate check,
   - searching for a customer,
   - creating a PDF slip or a print preview.
5. Use the slicers on the **Dashboard** sheet and press **Refresh Dashboard** after adding new data.

PDF slips are saved in `Documents\Appointment_Slip` on your computer.

To read the code, press `Alt + F11` to open the VBA editor and open **BookingForm**.

---

## Project Structure

| Sheet | Purpose |
|---|---|
| Customer_Master | Customer details |
| Appointment_Booking | All bookings (data saved by the form) |
| Payment_Entry | Payment records |
| charts | PivotTables and chart data |
| Dashboard | KPIs, charts and slicers |
| Reports | Summary reports |

*(Sheet names may vary slightly in the workbook.)*

---

## Screenshots

*Add screenshots of the Dashboard and the Booking Form here.*

---

## Author

**Mandeep Kaur**
