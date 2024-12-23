Office.onReady(() => {
    console.log("Office.js is ready.");

    document.addEventListener("DOMContentLoaded", async () => {
        console.log("DOM fully loaded and parsed.");

        const galleryDiv = document.getElementById("wallpaper-gallery");
        const statusDiv = document.getElementById("status");

        // Check for DOM elements
        if (!galleryDiv || !statusDiv) {
            console.error("Required DOM elements not found!");
            return;
        }

        console.log("Both DOM elements found. Proceeding...");

        // Fetch images from Pexels API
        async function fetchWallpapers() {
            const apiKey = "tiHzZylNhvmXcYiLM2zNGmUXO5m1hfxGCD0zyg44r74XbXhi0govsIqM"; // Replace with your Pexels API key
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
                }));
            } catch (error) {
                console.error("Error fetching wallpapers from Pexels:", error);
                return [];
            }
        }

        // Render images in the gallery
        function renderGallery(photos) {
            console.log("Rendering photos:", photos);

            photos.forEach((photo, index) => {
                if (!photo || !photo.urls || !photo.urls.small) {
                    console.warn("Invalid photo object:", photo);
                    return;
                }

                // Create image element
                const img = document.createElement("img");
                img.src = photo.urls.small;
                img.alt = `Image ${index + 1}`;
                img.style.cursor = "pointer";

                img.addEventListener("click", () => {
                    const option = prompt(
                        "Choose where to insert the image:\n1. Background\n2. Header\n3. Footer",
                        "1"
                    );
                    if (option === "1") {
                        setDocumentBackground(photo);
                    } else if (option === "2") {
                        setHeaderImage(photo);
                    } else if (option === "3") {
                        setFooterImage(photo);
                    } else {
                        alert("Invalid option. Please choose 1, 2, or 3.");
                    }
                });

                galleryDiv.appendChild(img);
            });

            if (photos.length === 0) {
                statusDiv.textContent = "No photos available.";
            }
        }

        // Set image as document background
        async function setDocumentBackground(photo) {
            try {
                const base64Image = await fetchImageAsBase64(photo.urls.regular);
                await Word.run(async (context) => {
                    const sections = context.document.sections;
                    sections.load("items");
                    await context.sync();

                    const body = sections.items[0].body;
                    body.insertInlinePictureFromBase64(
                        base64Image,
                        Word.InsertLocation.replace
                    );
                    await context.sync();
                });
                alert("Image set as document background!");
            } catch (error) {
                console.error("Error setting background:", error);
            }
        }

        // Set image in the header
        async function setHeaderImage(photo) {
            try {
                const base64Image = await fetchImageAsBase64(photo.urls.regular);
                await Word.run(async (context) => {
                    const sections = context.document.sections;
                    sections.load("items");
                    await context.sync();

                    const header = sections.items[0].getHeader("primary");
                    header.insertInlinePictureFromBase64(
                        base64Image,
                        Word.InsertLocation.replace
                    );
                    await context.sync();
                });
                alert("Image set in the header!");
            } catch (error) {
                console.error("Error setting header image:", error);
            }
        }

        // Set image in the footer
        async function setFooterImage(photo) {
            try {
                const base64Image = await fetchImageAsBase64(photo.urls.regular);
                await Word.run(async (context) => {
                    const sections = context.document.sections;
                    sections.load("items");
                    await context.sync();

                    const footer = sections.items[0].getFooter("primary");
                    footer.insertInlinePictureFromBase64(
                        base64Image,
                        Word.InsertLocation.replace
                    );
                    await context.sync();
                });
                alert("Image set in the footer!");
            } catch (error) {
                console.error("Error setting footer image:", error);
            }
        }

        // Fetch and convert image to Base64
        async function fetchImageAsBase64(imageUrl) {
            try {
                const response = await fetch(imageUrl);
                const blob = await response.blob();
                return new Promise((resolve, reject) => {
                    const reader = new FileReader();
                    reader.onloadend = () =>
                        resolve(reader.result.split(",")[1]); // Extract Base64
                    reader.onerror = reject;
                    reader.readAsDataURL(blob);
                });
            } catch (error) {
                console.error("Error fetching or converting image to Base64:", error);
                return null;
            }
        }

        // Fetch and display images
        const photos = await fetchWallpapers();
        renderGallery(photos);
    });
});
