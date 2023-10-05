async function melonSearch(query) {
  return new Promise((resolve, reject) => {
    const url = "https://www.melon.com/search/keyword/index.json";
    const jscallback = `jQuery${Math.floor(Math.random() * 1000000000)}_${Date.now()}`;

    const script = document.createElement('script');
    script.src = `${url}?jscallback=${jscallback}&query=${query}`;

    window[jscallback] = function(data) {
      resolve(data);
      delete window[jscallback];
    };

    script.onerror = () => {
      reject(new Error("script error"));
    };

    document.body.appendChild(script);
    script.onload = () => {
      document.body.removeChild(script);
    };
  });
}

