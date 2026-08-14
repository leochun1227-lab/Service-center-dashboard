window.serviceCentreData = {
  meta: {
    lastUpdated: "10 Aug 2026, 2:15 PM",
    months: ["Jul 2026", "Jun 2026", "May 2026"],
    yards: ["All Yards", "Factory Service", "Melbourne", "Sydney", "Brisbane"],
    invoiceScopes: ["Internal + External", "External Only", "Internal Only"]
  },
  pages: {
    overview: {
      title: "Service Centre Overview",
      subtitle: "C4C and SAP adoption snapshot for dealer yards and factory service.",
      kpis: [
        { title: "New Tickets", value: "426", detail: "$1.84M estimated ticket value across all yards", icon: "NT", tone: "blue", badge: "+8.4%", badgeTone: "up" },
        { title: "Invoiced Tickets", value: "318", detail: "$1.29M posted with internal and external split", icon: "IV", tone: "green", badge: "+5.1%", badgeTone: "up" },
        { title: "Open Tickets", value: "198", detail: "$742K still open, mostly factory and Melbourne", icon: "OP", tone: "orange", badge: "Watch", badgeTone: "warn" },
        { title: "Invoiced Hours", value: "3,482", detail: "Booked hours tracked against monthly payroll hours", icon: "HR", tone: "teal", badge: "92.6%", badgeTone: "up" },
        { title: "Active Yards", value: "82%", detail: "9 of 11 yards creating and invoicing tickets consistently", icon: "AD", tone: "blue", badge: "9 / 11", badgeTone: "up" },
        { title: "Data Exceptions", value: "14", detail: "Missing invoice type, yard mapping or technician assignment", icon: "DQ", tone: "red", badge: "14 gaps", badgeTone: "down" }
      ],
      yardActivity: [
        { yard: "Factory", color: "#1f6feb", newTickets: 118, invoiced: 92, open: 57 },
        { yard: "Melbourne", color: "#17a6ad", newTickets: 142, invoiced: 111, open: 66 },
        { yard: "Sydney", color: "#f58b1f", newTickets: 101, invoiced: 78, open: 43 },
        { yard: "Brisbane", color: "#7c3aed", newTickets: 65, invoiced: 37, open: 32 }
      ],
      invoiceMix: {
        total: "$1.29M",
        segments: [
          { name: "External", percent: 72, amount: "$0.93M", color: "#1f6feb" },
          { name: "Internal", percent: 28, amount: "$0.36M", color: "#17a6ad" }
        ]
      },
      yardSummary: [
        { yard: "Factory Service", color: "#1f6feb", newTickets: 118, invoiced: 92, open: 57, hours: 1084, status: "Backlog rising", statusTone: "watch" },
        { yard: "Melbourne", color: "#17a6ad", newTickets: 142, invoiced: 111, open: 66, hours: 1212, status: "Healthy flow", statusTone: "good" },
        { yard: "Sydney", color: "#f58b1f", newTickets: 101, invoiced: 78, open: 43, hours: 734, status: "Stable", statusTone: "good" },
        { yard: "Brisbane", color: "#7c3aed", newTickets: 65, invoiced: 37, open: 32, hours: 452, status: "Low adoption", statusTone: "risk" }
      ],
      utilisation: [
        { yard: "Factory", ratio: 91, note: "Assigned hours are close to payroll capacity", tone: "good" },
        { yard: "Melbourne", ratio: 96, note: "Strong usage, near full load", tone: "good" },
        { yard: "Sydney", ratio: 84, note: "Reasonable headroom still available", tone: "watch" },
        { yard: "Brisbane", ratio: 67, note: "Low usage suggests adoption or scheduling gaps", tone: "risk" }
      ],
      ageing: [
        { yard: "Factory", fresh: 22, warm: 18, stale: 17 },
        { yard: "Melbourne", fresh: 18, warm: 17, stale: 15 },
        { yard: "Sydney", fresh: 13, warm: 14, stale: 10 },
        { yard: "Brisbane", fresh: 9, warm: 12, stale: 11 }
      ],
      lifecycle: [
        { label: "Created", value: "426", detail: "Tickets opened in month from C4C" },
        { label: "Assigned", value: "401", detail: "Most tickets have yard and technician linked" },
        { label: "Invoiced", value: "318", detail: "Subset posted and available for finance alignment" },
        { label: "Still Open", value: "198", detail: "Open backlog requiring follow-up by yard" }
      ]
    },
    activity: {
      title: "Ticket Activity",
      subtitle: "New, invoiced and open movement over recent months.",
      kpis: [
        { title: "New Qty", value: "426", detail: "Current month ticket creation", icon: "NQ", tone: "blue", badge: "+8.4%", badgeTone: "up" },
        { title: "New Amount", value: "$1.84M", detail: "Estimated ticket value created", icon: "NA", tone: "teal", badge: "+6.2%", badgeTone: "up" },
        { title: "Invoiced Qty", value: "318", detail: "Tickets closed through invoicing", icon: "IQ", tone: "green", badge: "+5.1%", badgeTone: "up" },
        { title: "Invoiced Amount", value: "$1.29M", detail: "Posted amount this month", icon: "IA", tone: "orange", badge: "72% ext", badgeTone: "warn" }
      ],
      trend: [
        { month: "May", newTickets: 372, invoiced: 281, open: 172 },
        { month: "Jun", newTickets: 398, invoiced: 294, open: 184 },
        { month: "Jul", newTickets: 426, invoiced: 318, open: 198 }
      ],
      breakdown: [
        { title: "External", amount: "$0.93M", count: "228 tickets", note: "This is the amount expected to reconcile to finance." },
        { title: "Internal", amount: "$0.36M", count: "90 tickets", note: "Useful to separate workshop workload from customer revenue." }
      ],
      insights: [
        { title: "Factory and Melbourne are the heaviest ticket creators", text: "That suggests the first release should default to a yard comparison view rather than a single total number." },
        { title: "External invoicing deserves extra visibility", text: "Tim and finance will likely care most about whether customer-facing invoiced amounts line up cleanly." },
        { title: "Open growth is not terrible, but it is rising", text: "This helps position backlog as a workflow issue rather than only a volume issue." }
      ]
    },
    backlog: {
      title: "Open Backlog",
      subtitle: "Month-end open tickets, ageing and blocking reasons.",
      kpis: [
        { title: "Month-end Open", value: "198", detail: "Tickets still open after month close", icon: "OB", tone: "orange", badge: "+7.6%", badgeTone: "warn" },
        { title: "Open Value", value: "$742K", detail: "Estimated value still not invoiced", icon: "OV", tone: "red", badge: "Risk", badgeTone: "down" },
        { title: "30+ Days", value: "53", detail: "Older backlog needing follow-up", icon: "30", tone: "red", badge: "53", badgeTone: "down" },
        { title: "Missing Data", value: "14", detail: "Exceptions blocking clean closure", icon: "MD", tone: "blue", badge: "14 gaps", badgeTone: "warn" }
      ],
      rows: [
        { yard: "Factory Service", open: 57, stale: 17, value: "$228K", risk: "High", tone: "risk" },
        { yard: "Melbourne", open: 66, stale: 15, value: "$252K", risk: "Watch", tone: "watch" },
        { yard: "Sydney", open: 43, stale: 10, value: "$154K", risk: "Moderate", tone: "good" },
        { yard: "Brisbane", open: 32, stale: 11, value: "$108K", risk: "High", tone: "risk" }
      ],
      issues: [
        { title: "Missing technician assignment", text: "Some tickets exist in C4C but are not properly tied to a technician, which weakens hours reporting." },
        { title: "Invoice type not classified", text: "Without internal or external tagging, invoice reporting becomes harder to reconcile and explain." },
        { title: "Yard mapping inconsistency", text: "A few records still need clean yard labels so the by-yard dashboard stays trustworthy." }
      ]
    },
    hours: {
      title: "Hours & Payroll",
      subtitle: "Technician list, payroll hours, assigned hours and invoiced hours.",
      kpis: [
        { title: "Payroll Hours", value: "3,760", detail: "Monthly payable hours from HR", icon: "PH", tone: "blue", badge: "HR", badgeTone: "up" },
        { title: "Assigned Hours", value: "3,428", detail: "Hours assigned through ticket work", icon: "AH", tone: "teal", badge: "91%", badgeTone: "up" },
        { title: "Invoiced Hours", value: "3,482", detail: "Hours posted through invoicing", icon: "IH", tone: "green", badge: "92.6%", badgeTone: "up" },
        { title: "Low Coverage Yards", value: "1", detail: "Yards under 75% assigned to payroll coverage", icon: "LC", tone: "red", badge: "Brisbane", badgeTone: "down" }
      ],
      rows: [
        { yard: "Factory Service", techs: 8, payroll: 1190, assigned: 1084, invoiced: 1028, coverage: "91%" },
        { yard: "Melbourne", techs: 9, payroll: 1260, assigned: 1212, invoiced: 1186, coverage: "96%" },
        { yard: "Sydney", techs: 6, payroll: 870, assigned: 734, invoiced: 796, coverage: "84%" },
        { yard: "Brisbane", techs: 4, payroll: 440, assigned: 398, invoiced: 472, coverage: "67%" }
      ],
      balance: [
        { yard: "Factory", ratio: 91, note: "Near target utilisation", tone: "good" },
        { yard: "Melbourne", ratio: 96, note: "Very strong utilisation", tone: "good" },
        { yard: "Sydney", ratio: 84, note: "Moderate headroom", tone: "watch" },
        { yard: "Brisbane", ratio: 67, note: "Likely usage or coding gap", tone: "risk" }
      ]
    }
  }
};
