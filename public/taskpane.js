Office.onReady(() => {
    console.log("Office.js is ready.");

    document.addEventListener("DOMContentLoaded", async () => {
        console.log("DOM fully loaded and parsed.");

        const galleryDiv = document.getElementById("wallpaper-gallery");
        const statusDiv = document.getElementById("status");

        // Check if DOM elements exist
        if (!galleryDiv || !statusDiv) {
            console.error("Required DOM elements not found!");
            return;
        }

        console.log("Both DOM elements found. Proceeding...");

        // Fetch images from Pexels API
        async function fetchWallpapers() {
            const apiKey = "YOUR_PEXELS_API_KEY"; // Replace with your Pexels API key
            const url = "https://api.pexels.com/v1/curated?per_page=10";

            try {
                console.log("Fetching wallpapers from Pexels...");
                const response = await fetch(url, {
                    headers: {
                        Authorization: apiKey,
                    },
                });

                const data = await response.json();
                console.log("Pexels API Response:", data);

                return data.photos.map((photo) => ({
                    urls: {
                        small: photo.src.small,
                        regular: photo.src.large,
                    },
                    description: photo.alt || "Pexels Image",
                    user: {
                        name: photo.photographer,
                        links: { html: photo.photographer_url },
                    },
                }));
            } catch (error) {
                console.error("Error fetching wallpapers from Pexels:", error);
                return [];
            }
        }

        // Render images in the gallery
        function renderGallery(photos) {
            console.log("Rendering photos:", photos);

            photos.forEach((photo) => {
                if (!photo || !photo.urls || !photo.urls.small) {
                    console.warn("Invalid photo object:", photo);
                    return;
                }

                // Create image element
                const img = document.createElement("img");
                img.src = photo.urls.small;
                img.alt = photo.description;
                img.style.cursor = "pointer";

                galleryDiv.appendChild(img);

                // Add attribution
                const attribution = document.createElement("p");
                attribution.innerHTML = `Photo by <a href="${photo.user.links.html}" target="_blank">${photo.user.name}</a>`;
                galleryDiv.appendChild(attribution);
            });

            if (photos.length === 0) {
                statusDiv.textContent = "No photos available.";
            }
        }

        // Fetch and display images
        const photos = await fetchWallpapers();
        renderGallery(photos);
    });
});
