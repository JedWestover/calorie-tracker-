import { useEffect, useMemo, useRef, useState } from "react";
import { useMsal } from "@azure/msal-react";
import { loginRequest } from "./authConfig";
import {
  LineChart,
  Line,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  Legend,
} from "recharts";

const GRAPH_BASE = ["https:", "", "graph.microsoft.com", "v1.0"].join("/");
const APP_VERSION = "1.1.0";
const USDA_API_KEY = import.meta.env.VITE_USDA_API_KEY || "";
const BACKUP_URL =
  GRAPH_BASE + "/me/drive/special/approot:/backup.json:/content";

const USDA_SEARCH_URL = USDA_API_KEY
  ? "https://api.nal.usda.gov/fdc/v1/foods/search/"
  : "/api/usda";

const OPEN_FOOD_FACTS_URL = "https://api.openfoodfacts.org/api/v2/search";

const ACTIVITY_DATABASE = [
  { name: "Walking", category: "Cardio", levels: { light: 2.8, moderate: 3.5, vigorous: 4.3 } },
  { name: "Running", category: "Cardio", levels: { light: 6.0, moderate: 9.8, vigorous: 12.0 } },
  { name: "Jogging", category: "Cardio", levels: { light: 5.0, moderate: 7.0, vigorous: 8.5 } },
  { name: "Cycling", category: "Cardio", levels: { light: 4.0, moderate: 6.8, vigorous: 10.0 } },
  { name: "Swimming", category: "Cardio", levels: { light: 5.8, moderate: 8.0, vigorous: 10.0 } },
  { name: "Hiking", category: "Outdoor", levels: { light: 4.0, moderate: 6.0, vigorous: 8.0 } },
  { name: "Strength Training", category: "Strength", levels: { light: 3.0, moderate: 5.0, vigorous: 6.0 } },
  { name: "Yoga", category: "Flexibility", levels: { light: 2.3, moderate: 3.0, vigorous: 4.0 } },
  { name: "Rowing", category: "Cardio", levels: { light: 4.8, moderate: 7.0, vigorous: 9.0 } },
  { name: "Elliptical", category: "Cardio", levels: { light: 4.0, moderate: 5.5, vigorous: 8.0 } },
  { name: "Jump Rope", category: "Cardio", levels: { light: 8.0, moderate: 10.0, vigorous: 12.3 } },
  { name: "Dancing", category: "Cardio", levels: { light: 3.5, moderate: 5.5, vigorous: 7.8 } },
  { name: "Basketball", category: "Sports", levels: { light: 4.5, moderate: 6.5, vigorous: 8.0 } },
  { name: "Soccer", category: "Sports", levels: { light: 5.0, moderate: 7.0, vigorous: 10.0 } },
  { name: "House Cleaning", category: "Daily Activity", levels: { light: 2.5, moderate: 3.5, vigorous: 4.5 } },
  { name: "Yard Work", category: "Daily Activity", levels: { light: 3.5, moderate: 5.0, vigorous: 6.5 } },
  { name: "HIIT", category: "Cardio", levels: { light: 6.0, moderate: 9.0, vigorous: 12.0 } },
];

function getLocalDateKey(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");

  return year + "-" + month + "-" + day;
}

function normalizeEntryDate(entry) {
  if (entry.date && /^\d{4}-\d{2}-\d{2}$/.test(entry.date)) {
    return entry.date;
  }

  const parsedDate = entry.date ? new Date(entry.date) : new Date();
  return Number.isNaN(parsedDate.getTime())
    ? getLocalDateKey(new Date())
    : getLocalDateKey(parsedDate);
}

function normalizeEntries(entries) {
  return entries.map(function (entry) {
    return { ...entry, date: normalizeEntryDate(entry) };
  });
}

function getOpenFoodFactsCalories(nutriments) {
  return roundOne(
    nutriments["energy-kcal_100g"] ||
    nutriments["energy-kcal_serving"] ||
    nutriments["energy-kcal"] ||
    0
  );
}

export default function App() {
  const { instance, accounts } = useMsal();

  const [goal, setGoal] = useState(function () {
    const saved = localStorage.getItem("goal");
    return saved ? Number(saved) : 2200;
  });

  const [bodyWeightLbs, setBodyWeightLbs] = useState(function () {
    const saved = localStorage.getItem("bodyWeightLbs");
    return saved ? Number(saved) : 180;
  });

  const [goalWeight, setGoalWeight] = useState(function () {
    const saved = localStorage.getItem("goalWeight");
    return saved ? Number(saved) : 170;
  });

  const [weightHistory, setWeightHistory] = useState(function () {
    const saved = localStorage.getItem("weightHistory");
    return saved ? JSON.parse(saved) : [];
  });

  const [foodName, setFoodName] = useState("");
  const [foodCalories, setFoodCalories] = useState("");
  const [foodProtein, setFoodProtein] = useState("");
  const [foodCarbs, setFoodCarbs] = useState("");
  const [foodFat, setFoodFat] = useState("");
  const [foodServingMultiplier, setFoodServingMultiplier] = useState("1");
  const [selectedFoodBase, setSelectedFoodBase] = useState(null);

  const [foodSearch, setFoodSearch] = useState("");
  const [foodResults, setFoodResults] = useState([]);
  const [foodSearchStatus, setFoodSearchStatus] = useState("");

  const [activitySearch, setActivitySearch] = useState("");
  const [activityResults, setActivityResults] = useState([]);
  const [selectedActivity, setSelectedActivity] = useState(null);
  const [activityIntensity, setActivityIntensity] = useState("moderate");
  const [activityDuration, setActivityDuration] = useState(30);
  const [activitySearchStatus, setActivitySearchStatus] = useState("");

  const [workoutName, setWorkoutName] = useState("");
  const [workoutCalories, setWorkoutCalories] = useState("");

  const [foodEntries, setFoodEntries] = useState(function () {
    const saved = localStorage.getItem("foodEntries");
    return saved ? normalizeEntries(JSON.parse(saved)) : [];
  });

  const [savedFoods, setSavedFoods] = useState(() => {
    const saved =
      localStorage.getItem("savedFoods");

    return saved
      ? JSON.parse(saved)
      : [];
  });

  const [workoutEntries, setWorkoutEntries] = useState(function () {
    const saved = localStorage.getItem("workoutEntries");
    return saved ? normalizeEntries(JSON.parse(saved)) : [];
  });

  const [todayKey, setTodayKey] = useState(() => getLocalDateKey(new Date()));

  const [activeTab, setActiveTab] = useState("food");
  const [status, setStatus] = useState("");
  const [autoBackupEnabled, setAutoBackupEnabled] = useState(true);
  const [autoRestoreEnabled, setAutoRestoreEnabled] = useState(true);
  const [lastBackup, setLastBackup] = useState("");
  const [lastRestore, setLastRestore] = useState("");

  const skipFirstBackup = useRef(true);
  const hasAutoRestored = useRef(false);

  const isSignedIn = accounts.length > 0;

  useEffect(function () {
    const timer = setInterval(function () {
      setTodayKey(getLocalDateKey(new Date()));
    }, 60000);

    return function () {
      clearInterval(timer);
    };
  }, []);

  useEffect(function () {
    localStorage.setItem("goal", String(goal));
  }, [goal]);

  useEffect(function () {
    localStorage.setItem("bodyWeightLbs", String(bodyWeightLbs));
  }, [bodyWeightLbs]);

  useEffect(function () {
    localStorage.setItem("goalWeight", String(goalWeight));
  }, [goalWeight]);

  useEffect(function () {
    localStorage.setItem("weightHistory", JSON.stringify(weightHistory));
  }, [weightHistory]);

  useEffect(function () {
    localStorage.setItem("foodEntries", JSON.stringify(foodEntries));
  }, [foodEntries]);

  useEffect(() => {
    localStorage.setItem(
      "savedFoods",
      JSON.stringify(savedFoods)
    );
  }, [savedFoods]);

  useEffect(function () {
    localStorage.setItem("workoutEntries", JSON.stringify(workoutEntries));
  }, [workoutEntries]);

  useEffect(function () {
    if (!isSignedIn) return;
    if (!autoRestoreEnabled) return;
    if (hasAutoRestored.current) return;

    hasAutoRestored.current = true;
    restoreFromOneDrive(true);
  }, [isSignedIn, autoRestoreEnabled]);

  useEffect(function () {
    if (!isSignedIn) return;
    if (!autoBackupEnabled) return;

    if (skipFirstBackup.current) {
      skipFirstBackup.current = false;
      return;
    }

    const timer = setTimeout(function () {
      backupToOneDrive(true);
    }, 5000);

    return function () {
      clearTimeout(timer);
    };
  }, [
    goal,
    bodyWeightLbs,
    weightHistory,
    foodEntries,
    workoutEntries,
    isSignedIn,
    autoBackupEnabled,
  ]);

  useEffect(function () {
    if (!selectedFoodBase) return;

    const multiplier = Number(foodServingMultiplier || 1);

    setFoodCalories(String(Math.round(selectedFoodBase.calories * multiplier)));
    setFoodProtein(String(roundOne(selectedFoodBase.protein * multiplier)));
    setFoodCarbs(String(roundOne(selectedFoodBase.carbs * multiplier)));
    setFoodFat(String(roundOne(selectedFoodBase.fat * multiplier)));
  }, [foodServingMultiplier, selectedFoodBase]);

  useEffect(function () {
    if (!selectedActivity) return;

    const calories = calculateActivityCalories(
      selectedActivity,
      activityIntensity,
      activityDuration,
      bodyWeightLbs
    );

    setWorkoutName(
      selectedActivity.name +
      " - " +
      capitalize(activityIntensity) +
      " - " +
      activityDuration +
      " min"
    );

    setWorkoutCalories(String(calories));
  }, [selectedActivity, activityIntensity, activityDuration, bodyWeightLbs]);

  async function signIn() {
    if (!instance) {
      setStatus("Sign-in requires localhost or an HTTPS connection.");
      return;
    }

    try {
      await instance.loginRedirect(loginRequest);
    } catch (error) {
      console.error("Login failed:", error);
      setStatus("Login failed. Check the console.");
    }
  }

  async function signOut() {
    try {
      await instance.logoutRedirect();
    } catch (error) {
      console.error("Logout failed:", error);
      setStatus("Logout failed. Check the console.");
    }
  }

  async function getAccessToken() {
    if (!accounts[0]) {
      throw new Error("No signed-in account.");
    }

    const tokenRequest = {
      scopes: loginRequest.scopes,
      account: accounts[0],
    };

    try {
      const result = await instance.acquireTokenSilent(tokenRequest);
      return result.accessToken;
    } catch (error) {
      console.error("Silent token failed:", error);
      await instance.acquireTokenRedirect(tokenRequest);
      throw error;
    }
  }

  async function graphFetch(url, options) {
    const token = await getAccessToken();

    const finalOptions = options || {};
    const finalHeaders = {
      Authorization: "Bearer " + token,
    };

    if (finalOptions.headers) {
      Object.assign(finalHeaders, finalOptions.headers);
    }

    return fetch(url, {
      method: finalOptions.method || "GET",
      headers: finalHeaders,
      body: finalOptions.body,
    });
  }

  async function testGraph() {
    try {
      setStatus("Testing Microsoft Graph...");

      const response = await graphFetch(GRAPH_BASE + "/me");

      if (!response.ok) {
        throw new Error("Graph test failed: " + response.status);
      }

      const data = await response.json();

      setStatus(
        "Graph connected. Hello, " +
        (data.displayName || data.userPrincipalName || "Microsoft user") +
        "!"
      );
    } catch (error) {
      console.error(error);
      setStatus("Graph test failed. Check the console.");
    }
  }

  async function backupToOneDrive(isAuto) {
    try {
      if (!isSignedIn) return;

      setStatus(isAuto ? "Auto-backup running..." : "Saving backup...");

      const backup = {
  app: "Calories Fitness Tracker",
  version: 7,

  goal: goal,
  bodyWeightLbs: bodyWeightLbs,
  goalWeight: goalWeight,

  weightHistory: weightHistory,
  foodEntries: foodEntries,
  workoutEntries: workoutEntries,

  savedFoods: savedFoods,

  lastUpdated: new Date().toISOString(),
};

      const response = await graphFetch(BACKUP_URL, {
        method: "PUT",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(backup, null, 2),
      });

      if (!response.ok) {
        throw new Error("Backup failed: " + response.status);
      }

      const time = new Date().toLocaleTimeString();
      setLastBackup(time);

      setStatus(
        isAuto
          ? "Auto-backup saved at " + time + "."
          : "Backup saved at " + time + "."
      );
    } catch (error) {
      console.error(error);
      setStatus("Backup failed. Check the console.");
    }
  }

  async function restoreFromOneDrive(isAuto) {
    try {
      setStatus(
        isAuto ? "Auto-restore checking OneDrive..." : "Restoring from OneDrive..."
      );

      const response = await graphFetch(BACKUP_URL);

      if (response.status === 404) {
        setStatus("No OneDrive backup found yet.");
        return;
      }

      if (!response.ok) {
        throw new Error("Restore failed: " + response.status);
      }

      const backup = await response.json();

      setGoal(Number(backup.goal || 2200));
      setBodyWeightLbs(Number(backup.bodyWeightLbs || 180));
      setWeightHistory(Array.isArray(backup.weightHistory) ? backup.weightHistory : []);
      setFoodEntries(
        Array.isArray(backup.foodEntries)
          ? normalizeEntries(backup.foodEntries)
          : []
      );
      setWorkoutEntries(
        Array.isArray(backup.workoutEntries)
          ? normalizeEntries(backup.workoutEntries)
          : []
      );
      setGoalWeight(
  Number(backup.goalWeight || 170)
);

setSavedFoods(
  Array.isArray(backup.savedFoods)
    ? backup.savedFoods
    : []
);

      const time = new Date().toLocaleTimeString();
      setLastRestore(time);

      setStatus(
        isAuto
          ? "Auto-restore completed at " + time + "."
          : "Backup restored from OneDrive at " + time + "."
      );
    } catch (error) {
      console.error(error);
      setStatus("Restore failed. Check the console.");
    }
  }

  function getNutrient(food, possibleNames) {
    const nutrients = Array.isArray(food.foodNutrients)
      ? food.foodNutrients
      : [];

    const match = nutrients.find(function (nutrient) {
      const nutrientName = String(nutrient.nutrientName || "").toLowerCase();

      return possibleNames.some(function (name) {
        return nutrientName.includes(name.toLowerCase());
      });
    });

    if (!match) return 0;

    return roundOne(Number(match.value || 0));
  }

  async function searchFoods() {
    const query = foodSearch.trim();
    if (!query) return;

    try {
      setFoodSearchStatus("Searching USDA and Open Food Facts...");
      setFoodResults([]);

      const searchUrl =
        USDA_SEARCH_URL +
        "?query=" +
        encodeURIComponent(query) +
        "&pageSize=10" +
        (USDA_API_KEY
          ? "&api_key=" + encodeURIComponent(USDA_API_KEY)
          : "");

      const openFoodFactsUrl =
        OPEN_FOOD_FACTS_URL +
        "?search_terms=" +
        encodeURIComponent(query) +
        "&page_size=10" +
        "&fields=code,product_name,brands,quantity,nutriments";

      const responses = await Promise.allSettled([
        fetch(searchUrl),
        fetch(openFoodFactsUrl),
      ]);

      const mappedResults = [];

      if (responses[0].status === "fulfilled" && responses[0].value.ok) {
        const data = await responses[0].value.json();
        const foods = Array.isArray(data.foods) ? data.foods : [];

        foods.forEach(function (food) {
          const item = {
            id: "usda-" + food.fdcId,
            source: "USDA",
            name: food.description || "Unnamed food",
            brand: food.brandOwner || food.brandName || food.dataType || "",
            calories: getNutrient(food, ["Energy"]),
            protein: getNutrient(food, ["Protein"]),
            carbs: getNutrient(food, ["Carbohydrate"]),
            fat: getNutrient(food, ["Total lipid", "fat"]),
            servingSize: "base amount from USDA result",
          };

          if (item.name && item.calories > 0) mappedResults.push(item);
        });
      }

      if (responses[0].status === "fulfilled" && !responses[0].value.ok) {
        console.error("USDA search failed:", responses[0].value.status);
      }

      if (responses[1].status === "fulfilled" && responses[1].value.ok) {
        const data = await responses[1].value.json();
        const products = Array.isArray(data.products) ? data.products : [];

        products.forEach(function (product) {
          const nutriments = product.nutriments || {};
          const item = {
            id: "off-" + (product.code || product.id || product.product_name),
            source: "Open Food Facts",
            name: product.product_name || product.product_name_en || "Unnamed food",
            brand: product.brands || "",
            calories: getOpenFoodFactsCalories(nutriments),
            protein: roundOne(nutriments.proteins_100g),
            carbs: roundOne(nutriments.carbohydrates_100g),
            fat: roundOne(nutriments.fat_100g),
            servingSize: product.quantity
              ? product.quantity + " (nutrition per 100g)"
              : "nutrition per 100g",
          };

          if (item.name && item.calories > 0) mappedResults.push(item);
        });
      }

      if (responses[1].status === "fulfilled" && !responses[1].value.ok) {
        console.error(
          "Open Food Facts search failed:",
          responses[1].value.status
        );
      }

      setFoodResults(mappedResults);
      setFoodSearch("");

      if (mappedResults.length === 0) {
        setFoodSearchStatus("No foods found with calorie data.");
      } else {
        setFoodSearchStatus(
          "Select a USDA or Open Food Facts result, then adjust serving multiplier."
        );
      }
    } catch (error) {
      console.error(error);
      setFoodSearchStatus("Food search failed. Check the console.");
    }
  }

function selectFoodResult(item) {
  setSelectedFoodBase(item);
  setFoodServingMultiplier("1");
  setFoodName(item.name);
  setFoodCalories(String(item.calories));
  setFoodProtein(String(item.protein || 0));
  setFoodCarbs(String(item.carbs || 0));
  setFoodFat(String(item.fat || 0));

setFoodSearchStatus(
  "✅ " +
    item.name +
    " selected (" +
    item.calories +
    " cal, " +
    item.protein +
    "g protein). Adjust serving size if needed, then click Add Food or Save Food."
);

  setFoodResults([]);
  setFoodSearch("");
}

  function searchActivities() {
    if (!activitySearch.trim()) return;

    const query = activitySearch.toLowerCase();

    const matches = ACTIVITY_DATABASE.filter(function (activity) {
      return (
        activity.name.toLowerCase().includes(query) ||
        activity.category.toLowerCase().includes(query)
      );
    }).slice(0, 10);

    setActivityResults(matches);

    if (matches.length === 0) {
      setActivitySearchStatus("No matching activities found.");
    } else {
      setActivitySearchStatus("Select an activity, intensity, and duration.");
    }
  }

function selectActivity(activity) {
  setSelectedActivity(activity);

  setActivitySearchStatus(
    "✅ " +
      activity.name +
      " selected. Adjust intensity and duration if needed, then click Add Workout."
  );

  setActivityResults([]);
  setActivitySearch("");
}
``

  function logWeight() {
    if (!bodyWeightLbs) return;

    const entry = {
      id: Date.now(),
      date: new Date().toLocaleDateString(),
      weight: Number(bodyWeightLbs),
    };

    setWeightHistory([entry].concat(weightHistory));
  }

  function deleteWeightEntry(id) {
    setWeightHistory(
      weightHistory.filter(function (entry) {
        return entry.id !== id;
      })
    );
  }

  function saveFoodTemplate() {
    if (!foodName || !foodCalories) return;

    const exists = savedFoods.some(function (food) {
      return food.name.toLowerCase() === foodName.toLowerCase();
    });

    if (exists) {
      alert("Food already exists in My Foods");
      return;
    }

    const food = {
      id: Date.now(),
      name: foodName,
      calories: Number(foodCalories),
      protein: Number(foodProtein || 0),
      carbs: Number(foodCarbs || 0),
      fat: Number(foodFat || 0),
    };

    setSavedFoods([food].concat(savedFoods));
  }

  function loadSavedFood(food) {
    setFoodName(food.name);
    setFoodCalories(String(food.calories));
    setFoodProtein(String(food.protein || 0));
    setFoodCarbs(String(food.carbs || 0));
    setFoodFat(String(food.fat || 0));
    setFoodServingMultiplier("1");
    setSelectedFoodBase(null);
  }

  function loadRecentWorkout(workout) {
  setWorkoutName(workout.name);
  setWorkoutCalories(String(workout.calories));

  if (workout.intensity) {
    setActivityIntensity(workout.intensity);
  }

  if (workout.duration) {
    setActivityDuration(workout.duration);
  }
}

  function addFood() {
  if (!foodName || !foodCalories) return;

  const newEntry = {
    id: Date.now(),
    date: todayKey,
    name: foodName,
    calories: Number(foodCalories),
    protein: Number(foodProtein || 0),
    carbs: Number(foodCarbs || 0),
    fat: Number(foodFat || 0),
    servingMultiplier: Number(foodServingMultiplier || 1),
  };

  setFoodEntries(foodEntries.concat(newEntry));

  setFoodName("");
  setFoodCalories("");
  setFoodProtein("");
  setFoodCarbs("");
  setFoodFat("");
  setFoodServingMultiplier("1");
  setSelectedFoodBase(null);
}

    function addWorkout() {
      if (!workoutName || !workoutCalories) return;

      const newEntry = {
        id: Date.now(),
        date: todayKey,
        name: workoutName,
        calories: Number(workoutCalories),
        intensity: activityIntensity,
        duration: Number(activityDuration || 0),
        bodyWeightLbs: Number(bodyWeightLbs || 0),
      };

      setWorkoutEntries(workoutEntries.concat(newEntry));

      setWorkoutName("");
      setWorkoutCalories("");
      setSelectedActivity(null);
      setActivityDuration(30);
      setActivityIntensity("moderate");
      setActivitySearch("");
setActivityResults([]);
setActivitySearchStatus("");

    }

    function deleteFood(id) {
      setFoodEntries(
        foodEntries.filter(function (item) {
          return item.id !== id;
        })
      );
    }

    function deleteSavedFood(id) {
  setSavedFoods(
    savedFoods.filter(function (food) {
      return food.id !== id;
    })
  );
}

    function deleteWorkout(id) {
      setWorkoutEntries(
        workoutEntries.filter(function (item) {
          return item.id !== id;
        })
      );
    }

    const todayFoodEntries = foodEntries.filter(function (entry) {
      return normalizeEntryDate(entry) === todayKey;
    });

    const todayWorkoutEntries = workoutEntries.filter(function (entry) {
      return normalizeEntryDate(entry) === todayKey;
    });

    const caloriesConsumed = useMemo(
      function () {
        return todayFoodEntries.reduce(function (sum, item) {
          return sum + Number(item.calories || 0);
        }, 0);
      },
      [todayFoodEntries]
    );

    const weightChartData = [...weightHistory]
      .reverse()
      .map((entry) => ({
        date: entry.date,
        weight: entry.weight,
        goal: goalWeight,
      }));


    const caloriesBurned = useMemo(
      function () {
        return todayWorkoutEntries.reduce(function (sum, item) {
          return sum + Number(item.calories || 0);
        }, 0);
      },
      [todayWorkoutEntries]
    );

    const totalProtein = useMemo(
      function () {
        return todayFoodEntries.reduce(function (sum, item) {
          return sum + Number(item.protein || 0);
        }, 0);
      },
      [todayFoodEntries]
    );

    const nutritionChartData = todayFoodEntries.map(
      function (entry) {
        return {
          name: entry.name,
          calories: Number(entry.calories || 0),
          protein: Number(entry.protein || 0),
          carbs: Number(entry.carbs || 0),
          fat: Number(entry.fat || 0),
        };
      }
    );

    const nutritionByDay = {};

    foodEntries.forEach(function (entry) {
      const day = entry.date || "Unknown";

      if (!nutritionByDay[day]) {
        nutritionByDay[day] = {
          date: day,
          calories: 0,
          protein: 0,
          carbs: 0,
          fat: 0,
        };
      }

      nutritionByDay[day].calories += Number(
        entry.calories || 0
      );

      nutritionByDay[day].protein += Number(
        entry.protein || 0
      );

      nutritionByDay[day].carbs += Number(
        entry.carbs || 0
      );

      nutritionByDay[day].fat += Number(
        entry.fat || 0
      );
    });

    const dailyNutritionData =
      Object.values(nutritionByDay);

    const dailyHistory = {};

    foodEntries.forEach(function (entry) {
      const day = normalizeEntryDate(entry);

      if (!dailyHistory[day]) {
        dailyHistory[day] = { date: day, consumed: 0, burned: 0 };
      }

      dailyHistory[day].consumed += Number(entry.calories || 0);
    });

    workoutEntries.forEach(function (entry) {
      const day = normalizeEntryDate(entry);

      if (!dailyHistory[day]) {
        dailyHistory[day] = { date: day, consumed: 0, burned: 0 };
      }

      dailyHistory[day].burned += Number(entry.calories || 0);
    });

    const previousDays = Object.values(dailyHistory)
      .filter(function (day) {
        return day.date !== todayKey;
      })
      .sort(function (a, b) {
        return b.date.localeCompare(a.date);
      });

      const recentFoods = [];

foodEntries.forEach(function (food) {
  const exists = recentFoods.some(function (item) {
    return item.name === food.name;
  });

  if (!exists) {
    recentFoods.push(food);
  }
});

const recentFoodsList =
  [...recentFoods]
    .sort((a, b) => b.id - a.id)
    .slice(0, 10);

    const foodUsageMap = {};

    const recentWorkouts = [];

workoutEntries.forEach(function (workout) {
  const exists = recentWorkouts.some(function (item) {
    return item.name === workout.name;
  });

  if (!exists) {
    recentWorkouts.push(workout);
  }
});

const recentWorkoutsList =
  [...recentWorkouts]
    .sort(function (a, b) {
      return b.id - a.id;
    })
    .slice(0, 10);

foodEntries.forEach(function (food) {
  const key = food.name;

  if (!foodUsageMap[key]) {
    foodUsageMap[key] = {
      ...food,
      count: 0,
    };
  }

  foodUsageMap[key].count += 1;
});

const mostUsedFoods = Object.values(foodUsageMap)
  .sort(function (a, b) {
    return b.count - a.count;
  })
  .slice(0, 10);

    const totalCarbs = useMemo(
      function () {
        return todayFoodEntries.reduce(function (sum, item) {
          return sum + Number(item.carbs || 0);
        }, 0);
      },
      [todayFoodEntries]
    );

    const totalFat = useMemo(
      function () {
        return todayFoodEntries.reduce(function (sum, item) {
          return sum + Number(item.fat || 0);
        }, 0);
      },
      [todayFoodEntries]
    );

    const currentWeight =
      weightHistory.length > 0 ? weightHistory[0].weight : bodyWeightLbs;

    const startingWeight =
      weightHistory.length > 0
        ? weightHistory[weightHistory.length - 1].weight
        : bodyWeightLbs;

    const totalWeightChange = roundOne(currentWeight - startingWeight);

    const poundsRemaining = roundOne(
      goalWeight - currentWeight
    );

    const weeksTracked =
      Math.max(weightHistory.length - 1, 1);

    const poundsPerWeek =
      roundOne(
        (startingWeight - currentWeight) /
        weeksTracked
      );

    let estimatedWeeksRemaining = null;

    if (
      poundsPerWeek > 0 &&
      currentWeight > goalWeight
    ) {
      estimatedWeeksRemaining =
        Math.ceil(
          (currentWeight - goalWeight) /
          poundsPerWeek
        );
    }

    const netCalories = caloriesConsumed - caloriesBurned;
    const remaining = goal - netCalories;

    const progress = Math.min(
      100,
      Math.max(0, Math.round((netCalories / Math.max(goal, 1)) * 100))
    );

    return (
      <div className="min-h-screen bg-slate-100 pb-24 text-slate-900">
        <header className="bg-gradient-to-r from-emerald-500 to-green-600 px-4 py-6 text-white">
          <div className="mx-auto max-w-4xl">
            <p className="text-sm opacity-90">USDA + Activity Search + Weekly Weight</p>
            <h1 className="text-3xl font-bold">Calories + Fitness</h1>
            <p className="mt-2 text-sm opacity-90">
              Track meals, workouts, macros, weight history, and OneDrive sync.
            </p>
          </div>
        </header>

        <main className="mx-auto max-w-4xl space-y-4 p-4">
          <section className="rounded-2xl bg-white p-4 shadow-sm">
            <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
              <div>
                <p className="font-semibold">Microsoft Account</p>
                {isSignedIn ? (
                  <p className="break-all text-sm text-green-700">
                    Signed in as {accounts[0].username}
                  </p>
                ) : (
                  <p className="text-sm text-slate-500">
                    Sign in to enable OneDrive backup and restore.
                  </p>
                )}
              </div>

              {isSignedIn ? (
                <button
                  onClick={signOut}
                  className="rounded-xl bg-red-600 px-4 py-3 font-bold text-white"
                >
                  Sign Out
                </button>
              ) : (
                <button
                  onClick={signIn}
                  className="rounded-xl bg-blue-600 px-4 py-3 font-bold text-white"
                >
                  Sign In
                </button>
              )}
            </div>
          </section>

          {isSignedIn && (
            <section className="rounded-2xl bg-white p-4 shadow-sm">
              <h2 className="text-xl font-bold">OneDrive Sync</h2>

              <div className="mt-4 grid gap-3 sm:grid-cols-3">
                <button
                  onClick={testGraph}
                  className="rounded-xl bg-slate-800 px-4 py-3 font-bold text-white"
                >
                  Test Graph Connection
                </button>

                <button
                  onClick={function () {
                    backupToOneDrive(false);
                  }}
                  className="rounded-xl bg-emerald-600 px-4 py-3 font-bold text-white"
                >
                  Backup Now
                </button>

                <button
                  onClick={function () {
                    restoreFromOneDrive(false);
                  }}
                  className="rounded-xl bg-indigo-600 px-4 py-3 font-bold text-white"
                >
                  Restore
                </button>
              </div>

              <label className="mt-4 flex items-center gap-3 rounded-xl bg-slate-50 p-3 text-sm">
                <input
                  type="checkbox"
                  checked={autoBackupEnabled}
                  onChange={function (event) {
                    setAutoBackupEnabled(event.target.checked);
                  }}
                />
                <span>Enable automatic OneDrive backup after changes</span>
              </label>

              <label className="mt-3 flex items-center gap-3 rounded-xl bg-slate-50 p-3 text-sm">
                <input
                  type="checkbox"
                  checked={autoRestoreEnabled}
                  onChange={function (event) {
                    setAutoRestoreEnabled(event.target.checked);
                  }}
                />
                <span>Enable automatic restore after sign-in</span>
              </label>

              <p className="mt-3 text-sm text-slate-500">
                Backup location: OneDrive app folder / backup.json
              </p>

              <p className="text-sm text-slate-500">
                Last backup: {lastBackup || "Never"}
              </p>

              <p className="text-sm text-slate-500">
                Last restore: {lastRestore || "Never"}
              </p>

              {status && (
                <p className="mt-3 rounded-xl bg-slate-50 p-3 text-sm">
                  {status}
                </p>
              )}
            </section>
          )}

          <div className="grid gap-4 sm:grid-cols-2">
            <section
              className={
                "rounded-2xl bg-white p-4 shadow-sm " +
                (activeTab === "food" ? "block" : "hidden") +
                " sm:block"
              }
            >
              <h2 className="text-xl font-bold">Food Tracking</h2>

              <div className="mt-4 space-y-3">
                <div className="rounded-xl bg-slate-50 p-3">
                  <label className="text-sm font-semibold text-slate-600">
                    Search food databases
                  </label>

                  <div className="mt-2 flex flex-col gap-2 sm:flex-row">
                    <input
                      className="w-full rounded-xl border px-4 py-3"
                      placeholder="Search food, e.g. banana or oatmeal"
                      value={foodSearch}
                      onChange={function (event) {
                        setFoodSearch(event.target.value);
                      }}
                    />

                    <button
                      onClick={searchFoods}
                      className="rounded-xl bg-slate-800 px-4 py-3 font-bold text-white"
                    >
                      Search
                    </button>
                  </div>

                  {foodSearchStatus && (
  <p className="mt-2 rounded-lg bg-emerald-50 p-2 text-sm text-emerald-700">
    {foodSearchStatus}
  </p>
)}
                </div>

                {foodResults.length > 0 && (
                  <div className="space-y-2">
                    {foodResults.map(function (item) {
                      return (
                        <button
                          key={item.id}
                          onClick={function () {
                            selectFoodResult(item);
                          }}
                          className="w-full rounded-xl border bg-white p-3 text-left hover:bg-emerald-50"
                        >
                          <p className="font-semibold">{item.name}</p>

                          <p className="text-xs font-semibold uppercase text-slate-400">
                            {item.source}
                          </p>

                          {item.brand && (
                            <p className="text-sm text-slate-500">
                              {item.brand}
                            </p>
                          )}

                          <p className="text-sm text-emerald-700">
                            Base: {item.calories} cal, {item.protein}g protein,{" "}
                            {item.carbs}g carbs, {item.fat}g fat
                          </p>

                          <p className="text-xs text-slate-400">
                            {item.servingSize}
                          </p>
                        </button>
                      );
                    })}
                  </div>
                )}

                <input
                  className="w-full rounded-xl border px-4 py-3"
                  placeholder="Food name"
                  value={foodName}
                  onChange={function (event) {
                    setFoodName(event.target.value);
                    setSelectedFoodBase(null);
                  }}
                />

                <div className="rounded-xl bg-emerald-50 p-3">
                  <label className="text-sm font-semibold text-emerald-800">
                    Serving multiplier
                  </label>

                  <input
                    className="mt-2 w-full rounded-xl border px-4 py-3"
                    placeholder="1 = base amount, 2 = double, 0.5 = half"
                    type="number"
                    inputMode="decimal"
                    step="0.25"
                    min="0"
                    value={foodServingMultiplier}
                    onChange={function (event) {
                      setFoodServingMultiplier(event.target.value);
                    }}
                  />

                  <p className="mt-2 text-xs text-emerald-700">
                    Example: use 0.5 for half serving, 2 for double serving.
                  </p>
                </div>

                <input
                  className="w-full rounded-xl border px-4 py-3"
                  placeholder="Calories"
                  type="number"
                  inputMode="numeric"
                  value={foodCalories}
                  onChange={function (event) {
                    setFoodCalories(event.target.value);
                    setSelectedFoodBase(null);
                  }}
                />

                <div className="grid grid-cols-3 gap-2">
                  <input
                    className="rounded-xl border px-3 py-3"
                    placeholder="Protein"
                    type="number"
                    inputMode="decimal"
                    value={foodProtein}
                    onChange={function (event) {
                      setFoodProtein(event.target.value);
                      setSelectedFoodBase(null);
                    }}
                  />

                  <input
                    className="rounded-xl border px-3 py-3"
                    placeholder="Carbs"
                    type="number"
                    inputMode="decimal"
                    value={foodCarbs}
                    onChange={function (event) {
                      setFoodCarbs(event.target.value);
                      setSelectedFoodBase(null);
                    }}
                  />

                  <input
                    className="rounded-xl border px-3 py-3"
                    placeholder="Fat"
                    type="number"
                    inputMode="decimal"
                    value={foodFat}
                    onChange={function (event) {
                      setFoodFat(event.target.value);
                      setSelectedFoodBase(null);
                    }}
                  />
                </div>

                <div className="grid grid-cols-2 gap-2">
                  <button
                    onClick={addFood}
                    className="rounded-xl bg-emerald-500 px-4 py-3 font-bold text-white"
                  >
                    Add Food
                  </button>

                  <button
                    onClick={saveFoodTemplate}
                    className="rounded-xl bg-indigo-600 px-4 py-3 font-bold text-white"
                  >
                    Save Food
                  </button>
                </div>
              </div>

              <div className="mt-5">
                <h3 className="font-semibold">
                  My Foods
                </h3>

                {savedFoods.length === 0 ? (
                  <p className="mt-2 text-sm text-slate-500">
                    No saved foods yet.
                  </p>
                ) : (
                  <div className="mt-3 space-y-2">
                    {savedFoods.map((food) => (
  <div
    key={food.id}
    className="flex items-center justify-between rounded-xl border p-3"
  >
    <button
      onClick={() => loadSavedFood(food)}
      className="flex-1 text-left"
    >
      <p className="font-semibold">
        {food.name}
      </p>

      <p className="text-sm text-slate-500">
        {food.calories} cal |{" "}
        {food.protein || 0}g protein |{" "}
        {food.carbs || 0}g carbs |{" "}
        {food.fat || 0}g fat
      </p>
    </button>

    <button
      onClick={() => deleteSavedFood(food.id)}
      className="ml-3 rounded-lg bg-red-50 px-3 py-2 text-sm font-semibold text-red-600"
    >
      Delete
    </button>
  </div>
))}

                  </div>
                )}
              </div>

<div className="mt-5">
  <h3 className="font-semibold">
    Recent Foods
  </h3>

  {recentFoodsList.length === 0 ? (
    <p className="mt-2 text-sm text-slate-500">
      No recent foods yet.
    </p>
  ) : (
    <div className="mt-3 space-y-2">
      {recentFoodsList.map((food) => (
        <button
          key={food.id}
          onClick={() => loadSavedFood(food)}
          className="w-full rounded-xl border p-3 text-left hover:bg-emerald-50"
        >
          <p className="font-semibold">
            {food.name}
          </p>

          <p className="text-sm text-slate-500">
            {food.calories} cal
          </p>
        </button>
      ))}
    </div>
  )}
</div>

<div className="mt-5">
  <h3 className="font-semibold">
    Most Used Foods
  </h3>

  {mostUsedFoods.length === 0 ? (
    <p className="mt-2 text-sm text-slate-500">
      No frequently used foods yet.
    </p>
  ) : (
    <div className="mt-3 space-y-2">
      {mostUsedFoods.map((food) => (
        <button
          key={food.name}
          onClick={() => loadSavedFood(food)}
          className="w-full rounded-xl border p-3 text-left hover:bg-blue-50"
        >
          <p className="font-semibold">
            {food.name}
          </p>

          <p className="text-sm text-slate-500">
            Used {food.count} times
          </p>

          <p className="text-xs text-slate-400">
            {food.calories} cal |{" "}
            {food.protein || 0}g protein
          </p>
        </button>
      ))}
    </div>
  )}
</div>

              <LogList
                title="Food Log"
                emptyText="No food entries yet."
                entries={todayFoodEntries}
                type="food"
                onDelete={deleteFood}
              />
            </section>

            <section
              className={
                "rounded-2xl bg-white p-4 shadow-sm " +
                (activeTab === "fitness" ? "block" : "hidden") +
                " sm:block"
              }
            >
              <h2 className="text-xl font-bold">Fitness Tracking</h2>

              <div className="mt-4 space-y-3">
                <div className="rounded-xl bg-slate-50 p-3">
                  <label className="text-sm font-semibold text-slate-600">
                    Search activity database
                  </label>

                  <div className="mt-2 flex flex-col gap-2 sm:flex-row">
                    <input
                      className="w-full rounded-xl border px-4 py-3"
                      placeholder="Search activity, e.g. walking or running"
                      value={activitySearch}
                      onChange={function (event) {
                        setActivitySearch(event.target.value);
                      }}
                    />

                    <button
                      onClick={searchActivities}
                      className="rounded-xl bg-slate-800 px-4 py-3 font-bold text-white"
                    >
                      Search
                    </button>
                  </div>

                  {activitySearchStatus && (
                    <p className="mt-2 text-sm text-slate-500">
                      {activitySearchStatus}
                    </p>
                  )}

                  {activityResults.length > 0 && (
                    <div className="mt-3 space-y-2">
                      {activityResults.map(function (activity) {
                        return (
                          <button
                            key={activity.name}
                            onClick={function () {
                              selectActivity(activity);
                            }}
                            className="w-full rounded-xl border bg-white p-3 text-left hover:bg-blue-50"
                          >
                            <p className="font-semibold">{activity.name}</p>
                            <p className="text-sm text-slate-500">
                              {activity.category}
                            </p>
                            <p className="text-sm text-blue-700">
                              Light {activity.levels.light} MET | Moderate{" "}
                              {activity.levels.moderate} MET | Vigorous{" "}
                              {activity.levels.vigorous} MET
                            </p>
                          </button>
                        );
                      })}
                    </div>
                  )}
                </div>

                <div className="rounded-xl bg-blue-50 p-3">
                  <label className="text-sm font-semibold text-blue-800">
                    Intensity
                  </label>

                  <select
                    className="mt-2 w-full rounded-xl border px-4 py-3"
                    value={activityIntensity}
                    onChange={function (event) {
                      setActivityIntensity(event.target.value);
                    }}
                  >
                    <option value="light">Light</option>
                    <option value="moderate">Moderate</option>
                    <option value="vigorous">Vigorous</option>
                  </select>

                  <label className="mt-3 block text-sm font-semibold text-blue-800">
                    Duration, minutes
                  </label>

                  <input
                    className="mt-2 w-full rounded-xl border px-4 py-3"
                    type="number"
                    inputMode="numeric"
                    value={activityDuration}
                    onChange={function (event) {
                      setActivityDuration(Number(event.target.value));
                    }}
                  />
                </div>

                <input
                  className="w-full rounded-xl border px-4 py-3"
                  placeholder="Workout name"
                  value={workoutName}
                  onChange={function (event) {
                    setWorkoutName(event.target.value);
                    setSelectedActivity(null);
                  }}
                />

                <input
                  className="w-full rounded-xl border px-4 py-3"
                  placeholder="Calories burned"
                  type="number"
                  inputMode="numeric"
                  value={workoutCalories}
                  onChange={function (event) {
                    setWorkoutCalories(event.target.value);
                    setSelectedActivity(null);
                  }}
                />

                <button
                  onClick={addWorkout}
                  className="w-full rounded-xl bg-blue-500 px-4 py-3 font-bold text-white"
                >
                  Add Workout
                </button>
              </div>

<div className="mt-5">
  <h3 className="font-semibold">
    Recent Workouts
  </h3>

  {recentWorkoutsList.length === 0 ? (
    <p className="mt-2 text-sm text-slate-500">
      No recent workouts yet.
    </p>
  ) : (
    <div className="mt-3 space-y-2">
      {recentWorkoutsList.map((workout) => (
        <button
          key={workout.id}
          onClick={() =>
            loadRecentWorkout(workout)
          }
          className="w-full rounded-xl border p-3 text-left hover:bg-blue-50"
        >
          <p className="font-semibold">
            {workout.name}
          </p>

          <p className="text-sm text-slate-500">
            {workout.calories} cal burned
          </p>

          <p className="text-xs text-slate-400">
            {workout.intensity || "manual"} |{" "}
            {workout.duration || 0} min
          </p>
        </button>
      ))}
    </div>
  )}
</div>

              <LogList
                title="Workout Log"
                emptyText="No workout entries yet."
                entries={todayWorkoutEntries}
                type="workout"
                onDelete={deleteWorkout}
              />
            </section>
          </div>

          <section className="grid gap-4 sm:grid-cols-3">
            <div className="rounded-2xl bg-white p-4 shadow-sm">
              <label className="text-sm font-semibold text-slate-600">
                Daily calorie goal
              </label>

              <input
                className="mt-2 w-full rounded-xl border border-slate-300 px-4 py-3"
                type="number"
                value={goal}
                onChange={(event) => {
                  setGoal(Number(event.target.value));
                }}
              />
            </div>

            <div className="rounded-2xl bg-white p-4 shadow-sm">
              <label className="text-sm font-semibold text-slate-600">
                Current Weight (lbs)
              </label>

              <input
                className="mt-2 w-full rounded-xl border border-slate-300 px-4 py-3"
                type="number"
                value={bodyWeightLbs}
                onChange={(event) => {
                  setBodyWeightLbs(Number(event.target.value));
                }}
              />

              <button
                onClick={logWeight}
                className="mt-3 w-full rounded-xl bg-indigo-600 px-4 py-3 font-bold text-white"
              >
                Log Current Weight
              </button>
            </div>

            <div className="rounded-2xl bg-white p-4 shadow-sm">
              <label className="text-sm font-semibold text-slate-600">
                Goal Weight (lbs)
              </label>

              <input
                className="mt-2 w-full rounded-xl border border-slate-300 px-4 py-3"
                type="number"
                value={goalWeight}
                onChange={(event) => {
                  setGoalWeight(Number(event.target.value));
                }}
              />
            </div>
          </section>

          <section className="grid grid-cols-2 gap-3 md:grid-cols-4 xl:grid-cols-8">
            <SummaryCard label="Consumed" value={Math.round(caloriesConsumed)} suffix="cal" />
            <SummaryCard label="Burned" value={Math.round(caloriesBurned)} suffix="cal" />
            <SummaryCard label="Net" value={Math.round(netCalories)} suffix="cal" />
            <SummaryCard label="Remaining" value={Math.round(remaining)} suffix="cal" />
            <SummaryCard label="Weight" value={Math.round(bodyWeightLbs)} suffix="lbs" />
            <SummaryCard label="Change" value={totalWeightChange} suffix="lbs" />
            <SummaryCard
              label="Goal Diff"
              value={poundsRemaining}
              suffix="lbs"
            />

            <SummaryCard
              label="Trend"
              value={poundsPerWeek}
              suffix="lb/wk"
            />

          </section>

          <SummaryCard
            label="Avg Calories"
            value={
              dailyNutritionData.length > 0
                ? Math.round(
                  caloriesConsumed /
                  dailyNutritionData.length
                )
                : 0
            }
            suffix="cal/day"
          />

          <section className="grid grid-cols-3 gap-3">
            <SummaryCard label="Protein" value={roundOne(totalProtein)} suffix="g" />
            <SummaryCard label="Carbs" value={roundOne(totalCarbs)} suffix="g" />
            <SummaryCard label="Fat" value={roundOne(totalFat)} suffix="g" />
          </section>

          <section className="rounded-2xl bg-white p-4 shadow-sm">
            <div className="mb-2 flex justify-between">
              <p className="font-semibold">Daily progress</p>
              <p className="text-sm text-slate-500">{progress}%</p>
            </div>

            <div className="h-4 overflow-hidden rounded-full bg-slate-200">
              <div
                className="h-full rounded-full bg-emerald-500"
                style={{ width: progress + "%" }}
              />
            </div>
          </section>

          <section className="rounded-2xl bg-white p-4 shadow-sm">
            <div className="flex items-center justify-between gap-3">
              <div>
                <h2 className="text-xl font-bold">Daily History</h2>
                <p className="text-sm text-slate-500">
                  Today, {todayKey}, is shown in the active logs above.
                </p>
              </div>
              <p className="text-sm font-semibold text-emerald-600">
                {todayFoodEntries.length + todayWorkoutEntries.length} today
              </p>
            </div>

            {previousDays.length === 0 ? (
              <p className="mt-4 text-sm text-slate-500">
                Previous days will appear here after you log another day.
              </p>
            ) : (
              <div className="mt-4 space-y-2">
                {previousDays.map(function (day) {
                  return (
                    <div
                      key={day.date}
                      className="flex flex-wrap items-center justify-between gap-2 rounded-xl border p-3"
                    >
                      <p className="font-semibold">{day.date}</p>
                      <p className="text-sm text-slate-600">
                        {Math.round(day.consumed)} consumed | {Math.round(day.burned)} burned | Net {Math.round(day.consumed - day.burned)} cal
                      </p>
                    </div>
                  );
                })}
              </div>
            )}
          </section>

          <section className="rounded-2xl bg-white p-4 shadow-sm">
            <h2 className="text-xl font-bold">
              Weight Progress
            </h2>

            {weightHistory.length === 0 ? (
              <p className="mt-3 text-slate-500">
                No weight entries yet.
              </p>
            ) : (
              <>
                <div className="mt-4 h-72">
                  <ResponsiveContainer width="100%" height="100%">
                    <LineChart data={weightChartData}>
                      <CartesianGrid strokeDasharray="3 3" />

                      <XAxis dataKey="date" />

                      <YAxis domain={["auto", "auto"]} />

                      <Tooltip />

                      <Line
                        type="monotone"
                        dataKey="weight"
                        stroke="#10b981"
                        strokeWidth={3}
                        dot={{ r: 4 }}
                      />

                      <Line
                        type="monotone"
                        dataKey="goal"
                        stroke="#ef4444"
                        strokeDasharray="5 5"
                        strokeWidth={2}
                      />

                    </LineChart>
                  </ResponsiveContainer>
                </div>

                {estimatedWeeksRemaining && (
                  <div className="mt-4 rounded-xl bg-green-50 p-4">
                    <p className="font-semibold text-green-800">
                      Estimated Goal Time
                    </p>

                    <p className="mt-1 text-green-700">
                      Approximately {estimatedWeeksRemaining} weeks remaining.
                    </p>
                  </div>
                )}

                <div className="mt-4 space-y-2">
                  {weightHistory.map((entry) => (
                    <div
                      key={entry.id}
                      className="flex items-center justify-between rounded-xl border p-3"
                    >
                      <div>
                        <p className="font-semibold">
                          {entry.date}
                        </p>

                        <p className="text-sm text-slate-500">
                          {entry.weight} lbs
                        </p>
                      </div>

                      <button
                        onClick={() =>
                          deleteWeightEntry(entry.id)
                        }
                        className="rounded-lg bg-red-50 px-3 py-2 text-sm font-semibold text-red-600"
                      >
                        Delete
                      </button>
                    </div>
                  ))}
                </div>
              </>
            )}
          </section>

          <section className="rounded-2xl bg-white p-4 shadow-sm">
            <h2 className="text-xl font-bold">
              Nutrition Dashboard
            </h2>

            {nutritionChartData.length === 0 ? (
              <p className="mt-3 text-slate-500">
                Add food entries to see nutrition trends.
              </p>
            ) : (
              <div className="mt-4 h-80">
                <ResponsiveContainer
                  width="100%"
                  height="100%"
                >
                  <BarChart
                    data={dailyNutritionData}
                  >
                    <CartesianGrid
                      strokeDasharray="3 3"
                    />

                    <XAxis dataKey="date" />

                    <YAxis />

                    <Tooltip />

                    <Legend />

                    <Bar
                      dataKey="protein"
                      fill="#10b981"
                    />

                    <Bar
                      dataKey="carbs"
                      fill="#3b82f6"
                    />

                    <Bar
                      dataKey="fat"
                      fill="#f59e0b"
                    />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            )}
          </section>

          <section className="rounded-2xl bg-white p-4 shadow-sm">
            <h2 className="text-xl font-bold">
              Calories Breakdown
            </h2>

            {nutritionChartData.length > 0 && (
              <div className="mt-4 h-80">
                <ResponsiveContainer
                  width="100%"
                  height="100%"
                >
                  <LineChart
                    data={nutritionChartData}
                  >
                    <CartesianGrid
                      strokeDasharray="3 3"
                    />

                    <XAxis dataKey="name" />

                    <YAxis />

                    <Tooltip />

                    <Line
                      type="monotone"
                      dataKey="calories"
                      stroke="#ef4444"
                      strokeWidth={3}
                    />
                  </LineChart>
                </ResponsiveContainer>
              </div>
            )}
          </section>

          <section className="rounded-2xl bg-white p-4 shadow-sm">
            <h2 className="text-xl font-bold">
              Daily Macronutrients
            </h2>

            <div className="mt-4 h-80">
              <ResponsiveContainer
                width="100%"
                height="100%"
              >
                <LineChart
                  data={dailyNutritionData}
                >
                  <CartesianGrid
                    strokeDasharray="3 3"
                  />

                  <XAxis dataKey="date" />

                  <YAxis />

                  <Tooltip />

                  <Legend />

                  <Line
                    dataKey="protein"
                    stroke="#10b981"
                  />

                  <Line
                    dataKey="carbs"
                    stroke="#3b82f6"
                  />

                  <Line
                    dataKey="fat"
                    stroke="#f59e0b"
                  />
                </LineChart>
              </ResponsiveContainer>
            </div>
          </section>

          <section className="rounded-2xl bg-white p-2 shadow-sm sm:hidden">
            <div className="grid grid-cols-2 gap-2">
              <TabButton
                label="Food"
                active={activeTab === "food"}
                onClick={function () {
                  setActiveTab("food");
                }}
                activeClass="bg-emerald-500 text-white"
              />

              <TabButton
                label="Fitness"
                active={activeTab === "fitness"}
                onClick={function () {
                  setActiveTab("fitness");
                }}
                activeClass="bg-blue-500 text-white"
              />
            </div>
          </section>

          <footer className="rounded-2xl bg-white p-4 text-center text-sm text-slate-500 shadow-sm">
  <p>
    Calories + Fitness Tracker
  </p>

  <p>
    Version {APP_VERSION}
  </p>
</footer>

        </main>

        <nav className="fixed bottom-0 left-0 right-0 border-t bg-white p-2 shadow-lg sm:hidden">
          <div className="mx-auto grid max-w-md grid-cols-2 gap-2">
            <TabButton
              label="Food"
              active={activeTab === "food"}
              onClick={function () {
                setActiveTab("food");
              }}
              activeClass="bg-emerald-500 text-white"
            />

            <TabButton
              label="Fitness"
              active={activeTab === "fitness"}
              onClick={function () {
                setActiveTab("fitness");
              }}
              activeClass="bg-blue-500 text-white"
            />
          </div>
        </nav>
      </div>
    );
  }

  function calculateActivityCalories(activity, intensity, durationMinutes, weightLbs) {
    if (!activity) return 0;

    const met = activity.levels[intensity] || activity.levels.moderate;
    const weightKg = Number(weightLbs || 0) * 0.453592;
    const minutes = Number(durationMinutes || 0);

    return Math.round((met * 3.5 * weightKg * minutes) / 200);
  }

  function capitalize(value) {
    if (!value) return "";
    return value.charAt(0).toUpperCase() + value.slice(1);
  }

  function roundOne(value) {
    return Math.round(Number(value || 0) * 10) / 10;
  }

  function SummaryCard(props) {
    return (
      <div className="rounded-2xl bg-white p-4 shadow-sm">
        <p className="text-xs font-semibold uppercase text-slate-500">
          {props.label}
        </p>
        <p className="mt-2 text-xl md:text-2xl font-bold text-slate-900">{props.value}</p>
        <p className="text-xs md:text-sm text-slate-400">{props.suffix}</p>
      </div>
    );
  }

  function TabButton(props) {
    return (
      <button
        onClick={props.onClick}
        className={
          "rounded-xl px-4 py-3 font-semibold " +
          (props.active ? props.activeClass : "bg-slate-100 text-slate-600")
        }
      >
        {props.label}
      </button>
    );
  }

  function LogList(props) {
    return (
      <div className="mt-5 space-y-3">
        <h3 className="font-semibold">{props.title}</h3>

        {props.entries.length === 0 ? (
          <p className="text-sm text-slate-500">{props.emptyText}</p>
        ) : (
          props.entries.map(function (entry) {
            return (
              <div
                key={entry.id}
                className="flex items-center justify-between rounded-xl border p-3"
              >
                <div>
                  <p className="font-semibold">{entry.name}</p>

                  {props.type === "food" ? (
                    <p className="text-sm text-slate-500">
                      {entry.calories} cal | {entry.protein || 0}g protein |{" "}
                      {entry.carbs || 0}g carbs | {entry.fat || 0}g fat | serving x{" "}
                      {entry.servingMultiplier || 1}
                    </p>
                  ) : (
                    <p className="text-sm text-slate-500">
                      {entry.calories} cal burned | {entry.intensity || "manual"} |{" "}
                      {entry.duration || 0} min
                    </p>
                  )}
                </div>

                <button
                  onClick={function () {
                    props.onDelete(entry.id);
                  }}
                  className="rounded-lg bg-red-50 px-3 py-2 text-sm font-semibold text-red-600"
                >
                  Delete
                </button>
              </div>
            );
          })
        )}
      </div>
    );
  }