# 🚗 Daily Car-Search Email — Setup Guide

This tool searches Canada-wide every morning for a **Mitsubishi Outlander PHEV (2023–2024)**
and emails you a ranked list of the best-priced listings, with an Excel spreadsheet attached.

It runs **100% free** on GitHub's servers. You do **not** need to install anything on your
computer, and you never leave it running. GitHub does all the work in the cloud every day.

**Total setup time: about 15 minutes.** Just follow the steps in order.

---

## What you'll need before starting

1. A **GitHub account** (free) — sign up at https://github.com if you don't have one.
2. A **Gmail account** (free) — this is the address the tool will send the daily email *from*.
   (It can email the results to any address you want — even the same Gmail.)

That's it.

---

## Step 1 — Copy the project into your own GitHub account ("Fork")

1. Open the project page: **[GitHub project link here](https://github.com/mdnaseershah/vehicle-search-automation)**
2. In the **top-right corner**, click the **"Fork"** button.
3. On the next screen, leave everything as-is and click **"Create fork"**.

You now have your **own private copy** of the whole tool. Everything from here happens in
*your* copy — you won't touch your friend's version.

---

## Step 2 — Create a Gmail "App Password" (so the tool can send email)

Google won't let a program use your normal Gmail password. Instead you create a special
16-character "App Password" just for this tool. It's safe and you can delete it anytime.

1. First, turn on **2-Step Verification** if you haven't already:
   go to https://myaccount.google.com/security → "2-Step Verification" → follow the steps.
   *(App Passwords only exist once 2-Step Verification is on.)*
2. Then open **https://myaccount.google.com/apppasswords**
3. Type a name like **`Car Search`** and click **Create**.
4. Google shows a **16-character password** (four groups of four letters).
   **Copy it now** and keep it handy — you'll paste it in the next step.
   *(Ignore the spaces; you can type it with or without them.)*

---

## Step 3 — Add your 3 secret settings to GitHub

These "secrets" tell the tool which email account to use. They are stored privately and
hidden — nobody, including people viewing your public project, can see them.

1. In **your forked copy** on GitHub, click the **"Settings"** tab (near the top).
2. In the left menu, click **"Secrets and variables"** → **"Actions"**.
3. Click the green **"New repository secret"** button and add these **three**, one at a time.
   For each one: type the **Name** *exactly* as shown, paste the **Value**, click **"Add secret"**.

   | Name (type exactly)  | Value (what to paste)                                        |
   |----------------------|--------------------------------------------------------------|
   | `GMAIL_ADDRESS`      | Your full Gmail address, e.g. `yourname@gmail.com`          |
   | `GMAIL_PASSWORD`     | The 16-character App Password from Step 2                    |
   | `RECIPIENT_EMAIL`    | Where you want the results sent (any email address)         |

   ⚠️ The names must match **exactly** (all capitals, with underscores). This is the one place
   where a typo will stop the emails, so double-check them.

---

## Step 4 — Turn on the automatic daily run

New forks have automation switched off for safety. Turn it on:

1. In your forked copy, click the **"Actions"** tab at the top.
2. If you see a message like *"Workflows aren't being run on this forked repository,"*
   click the green **"I understand my workflows, go ahead and enable them"** button.

Done — from now on it runs by itself every morning.

---

## Step 5 — Send yourself a test email right now (don't wait until tomorrow)

You don't have to wait for the morning to check it works:

1. Still on the **"Actions"** tab, click the workflow name on the left
   (it's called something like **"Vehicle Search"**).
2. On the right, click the **"Run workflow"** dropdown button, then the green
   **"Run workflow"** button inside it.
3. Wait about **2–4 minutes**. The line will turn **green** with a checkmark when it finishes.
4. Check the inbox of your `RECIPIENT_EMAIL` — the results email should be there.
   *(If it's not in the inbox, check the Spam/Promotions folder the first time and mark it "Not spam.")*

If the email arrives, **you're all set.** 🎉

---

## What happens from now on

- Every morning (**around 8:30 AM Eastern time**) you'll automatically get an email with the
  best current listings, ranked cheapest-first, plus an Excel file with more detail.
- You don't have to do anything. It keeps running for free, forever, on its own.
- Listings you've seen before are tracked, so you'll notice **price drops** and **new** cars,
  and there's a "Recently Sold / Removed" section for cars that have disappeared.

---

## If something goes wrong

- **No email at all after a test run:** Re-check Step 3 — the three secret **names** must be
  spelled exactly `GMAIL_ADDRESS`, `GMAIL_PASSWORD`, `RECIPIENT_EMAIL`, and the password must be
  the **App Password** from Step 2 (not your normal Gmail login password).
- **The test run shows a red X:** Click into it to see the message, but 9 times out of 10 it's
  one of the three secrets being missing or misspelled.
- **First email went to Spam:** Open it, click "Not spam," and future ones land in the inbox.

---

## Want to search for a *different* car or price?

The search settings (which car, max price, max mileage, the years) live inside the file
**`vehicle_search_automation.py`**. Changing those is a light bit of editing — if your friend
wants a different vehicle, that's the file to adjust, and it's worth asking whoever set this up
to help make that first change.
