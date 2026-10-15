# Nice Guy CRM: Setup Guide

This takes about 45 to 60 minutes the first time. You need:

- A **Supabase** account (stores your customers, free)
- A **Microsoft** account (you probably already have one, for Outlook; free)
- A **GitHub** account (hosts the app so your phone can open it, free)

Nothing here costs money on the free tiers.

Keep a notepad open. You will copy three values into `config.js`:

| Value | Where it comes from |
|---|---|
| Supabase Project URL | Step 1.4 |
| Supabase anon (public) key | Step 1.4 |
| Azure Application (client) ID | Step 2.4 |

---

## Part 1: Database (Supabase)

### 1.1 Create the account and project
1. Go to **https://supabase.com** and click **Start your project**. Sign up with Google or email.
2. Click **New project**.
3. Fill in:
   - **Name:** `nice-guy-crm`
   - **Database password:** click *Generate*, then save it in a password manager. You will rarely need it.
   - **Region:** pick the closest to Florida or your area, if available.
4. Click **Create new project** and wait about 2 minutes.

### 1.2 Create the tables and security rules
1. In the left menu, click **SQL Editor**, then **+ New query**.
2. Open `supabase-schema.sql` from the app folder, copy everything, and paste it into the editor.
3. Click **Run** (bottom right). You should see "Success. No rows returned."

### 1.3 Turn off public sign-ups
Only you should be able to log in.

1. Left menu: **Authentication**, then **Sign In / Providers** (on some screens **Providers** > **Email**).
2. Turn **off** "Allow new users to sign up". Save.

### 1.4 Create your login and copy the keys
1. Left menu: **Authentication**, then **Users**, then **Add user** > **Create new user**.
2. Enter your email and a password you will remember. Tick **Auto Confirm User**. Click **Create user**.
3. Go to **Project Settings** (gear icon, bottom left), then **API** (under Configuration).
4. Copy:
   - **Project URL** (looks like `https://abcd1234.supabase.co`)
   - **anon / public** key (a long string starting with `eyJ...`, or labelled "Publishable key" on newer screens)

   Do **not** copy the `service_role` key. It bypasses the security rules and must never go in the app.

Paste the URL and key into `config.js`, replacing the `PASTE_...` text in the first two lines.

---

## Part 2: Outlook connection (Microsoft Azure)

### 2.1 Open the Azure portal
1. Go to **https://portal.azure.com** and sign in with your Microsoft account. A personal Outlook.com, Hotmail, or Microsoft 365 account all work.
2. If it asks you to set up a directory or subscription, accept the defaults. It is free.

### 2.2 Register the app
1. Search for **Microsoft Entra ID** (formerly Azure Active Directory) and open it.
2. Left menu: **App registrations**, then **+ New registration**.
3. Fill in:
   - **Name:** `Nice Guy CRM`
   - **Supported account types:** choose **Accounts in any organizational directory and personal Microsoft accounts**. This is the option that works for both personal Outlook and work email.
   - **Redirect URI:** choose **Single-page application (SPA)** from the dropdown, and enter the address where the app will be hosted (you will know it after Part 3, so for now type a placeholder such as `https://yourname.github.io/nice-guy-crm/`).

     **It must end with a `/` and match exactly.** You can change it later under **Authentication**.
4. Click **Register**. On the overview page, copy **Application (client) ID**. This is the `MS_CLIENT_ID` value for `config.js`.

### 2.3 Permission to write to your calendar
1. In your new app, left menu: **API permissions**.
2. Click **+ Add a permission**, then **Microsoft Graph**, then **Delegated permissions**.
3. Search for `Calendars.ReadWrite`, tick it, and click **Add permissions**.
4. If you see a "Grant admin consent" button, you do not need it for a personal account. If you use a work account and the button is there, ask your IT person to click it.

### 2.4 Confirm the redirect page
1. Left menu: **Authentication**.
2. Under **Single-page application** you should see your URL. Make sure **Access tokens** and **ID tokens** are not required. The defaults are fine.

Paste the Application (client) ID into `config.js` as `MS_CLIENT_ID`.

---

## Part 3: Host the app (GitHub Pages)

Your iPhone can only use the app over a web address (HTTPS), so it has to be hosted. GitHub Pages does this for free.

### 3.1 Create the repository
1. Sign up or sign in at **https://github.com**.
2. Click **+** (top right), then **New repository**.
3. Name it `nice-guy-crm`. Set it to **Public**. Click **Create repository**.

   Public is required for free GitHub Pages. This is safe: the app contains no secret keys. Your customer data lives in Supabase and is locked to your login.

### 3.2 Upload the files
1. On the empty repository page, click **uploading an existing file**.
2. Drag in **everything** from the `nice-guy-crm` folder: `index.html`, `styles.css`, `app.js`, `config.js` (with your values filled in), `manifest.json`, `sw.js`, the `icons` folder, and `supabase-schema.sql`. Do not include `SETUP.md` if you want to keep it private, though it is harmless.
3. Click **Commit changes**.

### 3.3 Turn on Pages
1. Go to **Settings** (repository tab), then **Pages** (left menu).
2. Under **Build and deployment**, set **Source** to **Deploy from a branch**. Set **Branch** to `main` and folder to `/ (root)`. Click **Save**.
3. Wait 1 to 2 minutes, then refresh. The page shows your address, which will be:

   `https://YOUR-GITHUB-USERNAME.github.io/nice-guy-crm/`

4. **Go back to Azure (Part 2.2 step 3 / Authentication)** and make sure the Redirect URI is exactly this address, with the trailing `/`. If it differs by even one character, Outlook sign-in will fail.

### 3.4 Test it
Open the address in a browser. You should see the sign-in screen. Sign in with the email and password from step 1.4.

---

## Part 4: Install on your devices

### Windows laptop
1. Open the app address in **Microsoft Edge** (or Chrome).
2. Click the **...** menu (top right), then **Apps**, then **Install this site as an app**.
3. Name it and click **Install**. It now opens in its own window and appears in your Start menu.

### iPhone
1. Open the app address in **Safari**. This must be Safari, because iPhones only allow home-screen installs from Safari.
2. Tap the **Share** button (square with arrow), then scroll and tap **Add to Home Screen**.
3. Tap **Add**. The "Nice Guy CRM" icon appears on your home screen.
4. Open it from the icon and sign in.

### Connect Outlook on each device
Outlook is connected separately on each device (the laptop and the phone). On each one, tap **Connect Outlook** at the top and sign in with your Microsoft account. You only need to do this once per device, though you may be asked again occasionally.

---

## Part 5: Your first run

1. Sign in.
2. Tap **Connect Outlook** and approve the permission (Microsoft will ask you to allow the app to read and write your calendar). Approve it.
3. Tap **+ Add** and create a test customer.
4. Open the customer, add a note, and add a follow-up with a date and time. Tap **Add follow-up and send to Outlook**.
5. Open your Outlook calendar. The event should appear at that time with a reminder 15 minutes before. Your phone's Outlook app will show it too.

When you're happy, delete the test customer (open the customer, **Delete customer**).

---

## Troubleshooting

| Problem | Fix |
|---|---|
| Sign-in screen says "Setup not finished" | `config.js` still has `PASTE_...` values. Fill in the Supabase URL and key, then re-upload the file to GitHub. |
| "Sign-in failed: Invalid login credentials" | Check the email and password. Make sure you ticked **Auto Confirm User** in step 1.4. |
| Customers do not load, or "permission denied" | The SQL from step 1.2 did not run fully. Run it again in the SQL Editor. |
| "Outlook setup is not finished" | The Azure client ID is missing from `config.js`. |
| Outlook sign-in page says "redirect URI mismatch" | The Redirect URI in Azure (Authentication) does not exactly match your GitHub Pages address. Fix it there, including the trailing `/`. |
| Outlook error 403 | `Calendars.ReadWrite` was not added in step 2.3, or a work account needs admin consent. |
| Changes to the app are not showing on your phone | Close the app fully and reopen it. If it still looks old, clear the site data in Safari settings, or delete and re-add the home-screen icon. |
| Supabase says the project is paused | Free projects can pause after a period of no use. Open the project in supabase.com and click **Restore**. |

---

## Notes

- **Backups:** Supabase's free tier does not keep automatic backups. Every few months, you can save a copy of your data: in Supabase, go to **Table Editor**, open a table, and use **Export**.
- **Password:** if you forget your login, use **Authentication** > **Users** in Supabase to reset it.
- **Costs:** Supabase free, GitHub Pages free, Azure and Outlook free for this use.
- **Changing the app later:** upload the updated files to the GitHub repository (replace the old ones). The app updates automatically the next time you open it.
