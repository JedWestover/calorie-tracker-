# Calories + Fitness Tracker

Version 1.1.0 of a React and Vite app for tracking food, workouts, nutrition, weight progress, and OneDrive backups.

## Features

- Log food with calories, protein, carbohydrates, fat, and serving multipliers
- Search USDA FoodData Central and Open Food Facts
- Search activities and estimate calories burned from MET values
- Track weight history and visualize nutrition and weight trends
- Save frequently used foods and workouts
- Back up and restore data through Microsoft OneDrive
- Installable Progressive Web App (PWA)

## Requirements

- Node.js 20 or newer
- A USDA FoodData Central API key for USDA search
- An Azure app registration configured for Microsoft account sign-in and OneDrive access

## Local setup

```bash
npm install
```

Copy `src/.env.example` to `src/.env` and add your USDA API key:

```env
VITE_USDA_API_KEY=your_usda_api_key
```

Start the development server:

```bash
npm run dev
```

Open the local URL shown by Vite. Microsoft sign-in requires `localhost` or an HTTPS URL. Add that URL as a redirect URI in the Azure app registration, and grant delegated permissions for `User.Read` and `Files.ReadWrite`.

For Vercel, add `USDA_API_KEY` as an environment variable and redeploy. Keep the key server-side; do not use a `VITE_` prefix for this variable.

## Scripts

```bash
npm run dev      # Start the development server
npm run build    # Create a production build
npm run preview  # Preview the production build locally
npm run lint     # Run ESLint
```

## Deployment note

The repository includes Vercel serverless functions for `/api/usda` and `/api/open-food-facts`, which avoid browser CORS restrictions. Vercel is required because GitHub Pages cannot run these API functions. Configure the Azure redirect URI for the final HTTPS domain before enabling sign-in there.

Do not commit `src/.env` or any API keys. The repository ignores `.env` files; `src/.env.example` contains placeholders only.
