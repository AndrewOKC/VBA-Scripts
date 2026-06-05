# Segments vs. Audiences: The BI Translation

In traditional BI tools (Power BI, Tableau) or general data science operations (SQL), data is usually filtered retroactively. If you write a `WHERE` clause or apply a dashboard slicer today, it instantly filters all the historical data you collected last year.

GA4 splits this concept into two distinct tools with very different data-collection rules:

| Feature       | The BI / SQL Equivalent                 | When Does Data Start?                                | Where Does It Live?                           | Looker Studio?                     |
| ------------- | --------------------------------------- | ---------------------------------------------------- | --------------------------------------------- | ---------------------------------- |
| **Segments**  | A `WHERE` clause on the event table     | **Retroactive.** Works instantly on historical data. | Explorations (Custom Reports) only.           | **No.** Confined to the GA4 UI.    |
| **Audiences** | A user-level flag (`is_registered = 1`) | **Forward-looking.** Starts at zero upon creation.   | Everywhere (Standard Reports & Explorations). | **Yes.** Available as a dimension. |

---

## GA4 Segments (The On-the-Fly Filter)

Segments are exploratory tools used to isolate specific subsets of your data to analyze UI performance, user flows, or A/B test results.

- **How they work:** They filter the underlying dataset on the fly based on the events that occurred _within your selected date range_.
- **The Lookback Reality:** Because Segments are retroactive, applying a Segment instantly gives you access to all historical data. If you set a time-windowed condition (e.g., users who did X within Y days), GA4 allows a maximum lookback window of **60 days**.
- **The Tooling Limitation:** Segments only exist inside the specific "Exploration" where you created them. They cannot be applied to standard out-of-the-box GA4 reports, and **they cannot be exported to Looker Studio.**

## GA4 Audiences (The State-Based Bucket)

Audiences are persistent groups of users who share a specific attribute or behavior state.

- **How they work:** When you create an Audience, GA4 essentially creates a new bucket and starts watching the live data stream. When a user meets the criteria, they are dropped into the bucket and stay there (up to your defined membership duration).
- **The Lookback Reality (The "Day Zero" Rule):** This is the biggest mental shift from Power BI. **Audiences are NOT retroactive for GA4 internal reporting.** They begin accumulating users from the exact moment you click "Save." If you create an Audience today, your historical data for it is zero.
    > _Note: If GA4 is linked to Google Ads, the Ads platform will pull in up to 30 days of historical data for remarketing. Since our focus is UX and UI analysis, this 30-day grace period does not apply to our GA4 reports._
- **The Advantage:** Because Audiences function as a persistent user dimension, they are available globally across all standard reports. Most importantly, **Audiences can be exported to Looker Studio** for external dashboarding.

---

## The Timeline Trap: January Registration vs. February Reporting

To understand the practical difference, let's look at how this plays out when analyzing the UX of our landing pages using a "Registered Users" group (users who fired `login` or `sign_up`).

**The Scenario:** You want to analyze UI performance for the month of February. A user registered on our site back in January. They return in February to use the site, but because their session was still active, they didn't need to log in again.

- **Using a Segment (Event-Bound):** If you apply your "Registered Users" Segment to a February Exploration, **that user's activity will not show up.** Why? Segments look for the triggering event _strictly within the report's date range_. Because they didn't fire the `login` or `sign_up` event during February, their February UI interactions are filtered out entirely.
- **Using an Audience (State-Bound):** If you built the "Registered Users" Audience back in January, the user was permanently flagged with that state. When you pull a February report using the Audience filter, **all of their February activity will be visible.** They carry their "Registered" status with them into future months, regardless of whether they fired the event again during the reporting window.
