Office.onReady(() => {
    console.log("Office.js is ready.");

    document.addEventListener("DOMContentLoaded", async () => {
        console.log("DOM fully loaded and parsed.");

        const galleryDiv = document.getElementById("wallpaper-gallery");
        const statusDiv = document.getElementById("status");
        const optionsDiv = document.getElementById("options");
        let selectedPhoto = null;

        // Check for required DOM elements
        if (!galleryDiv || !statusDiv || !optionsDiv) {
            console.error("Required DOM elements not found!");
            return;
        }

        console.log("All required DOM elements found. Proceeding...");

        // Fetch wallpapers from Pexels
async function fetchWallpapers() {
    const apiKey = "tiHzZylNhvmXcYiLM2zNGmUXO5m1hfxGCD0zyg44r74XbXhi0govsIqM"; // Replace with your actual Pexels API Key
    const url = "https://api.pexels.com/v1/curated?per_page=10";

    try {
        console.log("Fetching wallpapers from Pexels...");
        const response = await fetch(url, {
            headers: {
                Authorization: apiKey, // Pexels API requires this in the header
            },
        });

        const data = await response.json();
        console.log("Pexels API Response:", data);

        // Transform the response to match the existing format
        const photos = data.photos.map((photo) => ({
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

        return photos;
    } catch (error) {
        console.error("Error fetching wallpapers from Pexels:", error);
        return [];
    }
}

        // Render wallpapers in the gallery
        function renderGallery(photos) {
            console.log("Rendering photos:", photos);

            photos.forEach((photo) => {
                if (!photo || !photo.urls || !photo.urls.small) {
                    console.warn("Invalid photo object:", photo);
                    return; // Skip invalid entries
                }

                // Create image element
                const img = document.createElement("img");
                img.src = photo.urls.small; // Use Unsplash image URL
                img.alt = photo.description || "Unsplash Image";
                img.style.cursor = "pointer";

                // Add click event to select the photo
                img.addEventListener("click", () => {
                    selectedPhoto = photo;
                    optionsDiv.style.display = "block"; // Show options
                    statusDiv.textContent =
                        "Image selected. Choose an option below.";
                    console.log("Selected photo:", selectedPhoto);
                });

                galleryDiv.appendChild(img);

                // Add attribution
                const attribution = document.createElement("p");
                attribution.innerHTML = `Photo by <a href="${photo.user.links.html}?utm_source=word_addin&utm_medium=referral" target="_blank">${photo.user.name}</a> on <a href="https://unsplash.com/?utm_source=word_addin&utm_medium=referral" target="_blank">Unsplash</a>`;
                galleryDiv.appendChild(attribution);
            });
        }

        // Set image as document background
        async function setDocumentBackground() {
            if (!selectedPhoto) {
                alert("No image selected.");
                return;
            }
            console.log("Selected photo for background:", selectedPhoto);

            try {
                const base64Image = await fetchImageAsBase64(selectedPhoto.urls.regular);
                console.log("Base64 Image:", base64Image);

                await Word.run(async (context) => {
                    console.log("Running Word API to insert background...");
                    const sections = context.document.sections;
                    sections.load("items");
                    await context.sync();

                    const body = sections.items[0].body;
                    body.insertInlinePictureFromBase64(
                        base64Image,
                        Word.InsertLocation.replace
                    );
                    await context.sync();
                    console.log("Background set successfully!");
                });
            } catch (error) {
                console.error("Error setting background:", error);
            }
        }

        // Set image in the header
        async function setHeaderImage() {
            if (!selectedPhoto) {
                alert("No image selected.");
                return;
            }
            console.log("Selected photo for header:", selectedPhoto);

            try {
                const base64Image = await fetchImageAsBase64(selectedPhoto.urls.regular);

                await Word.run(async (context) => {
                    console.log("Running Word API to insert header...");
                    const sections = context.document.sections;
                    sections.load("items");
                    await context.sync();

                    const header = sections.items[0].getHeader("primary");
                    header.insertInlinePictureFromBase64(
                        base64Image,
                        Word.InsertLocation.replace
                    );
                    await context.sync();
                    console.log("Header image set successfully!");
                });
            } catch (error) {
                console.error("Error setting header image:", error);
            }
        }

        // Set image in the footer
        async function setFooterImage() {
            if (!selectedPhoto) {
                alert("No image selected.");
                return;
            }
            console.log("Selected photo for footer:", selectedPhoto);

            try {
                const base64Image = await fetchImageAsBase64(selectedPhoto.urls.regular);

                await Word.run(async (context) => {
                    console.log("Running Word API to insert footer...");
                    const sections = context.document.sections;
                    sections.load("items");
                    await context.sync();

                    const footer = sections.items[0].getFooter("primary");
                    footer.insertInlinePictureFromBase64(
                        base64Image,
                        Word.InsertLocation.replace
                    );
                    await context.sync();
                    console.log("Footer image set successfully!");
                });
            } catch (error) {
                console.error("Error setting footer image:", error);
            }
        }

        // Fetch and convert image to Base64
        async function fetchImageAsBase64(imageUrl) {
            console.log("Fetching image:", imageUrl);

            try {
                const response = await fetch(imageUrl);
                console.log("Image fetched successfully:", response);

                const blob = await response.blob();
                console.log("Blob created from response:", blob);

                return new Promise((resolve, reject) => {
                    const reader = new FileReader();
                    reader.onloadend = () => {
                        console.log("Base64 Conversion Complete:", reader.result);
                        resolve(reader.result.split(",")[1]); // Extract Base64
                    };
                    reader.onerror = reject;
                    reader.readAsDataURL(blob);
                });
            } catch (error) {
                console.error("Error fetching or converting image to Base64:", error);
                return null;
            }
        }

        // Add event listeners for options
        document
            .getElementById("set-background")
            .addEventListener("click", () => setDocumentBackground());
        document
            .getElementById("set-header")
            .addEventListener("click", () => setHeaderImage());
        document
            .getElementById("set-footer")
            .addEventListener("click", () => setFooterImage());

        // Load wallpapers and render gallery
        const photos = await fetchWallpapers();
        if (photos.length > 0) {
            renderGallery(photos);
        } else {
            statusDiv.textContent = "No photos found. Please try again.";
        }
    });
});
