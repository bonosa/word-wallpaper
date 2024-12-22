Office.onReady(() => {
    console.log("Office.js is ready.");

    document.addEventListener("DOMContentLoaded", async () => {
        console.log("DOM fully loaded and parsed.");

        const galleryDiv = document.getElementById("wallpaper-gallery");
        const statusDiv = document.getElementById("status");
        const optionsDiv = document.getElementById("options");
        let selectedPhoto = null;

        // Check for required DOM elements
        if (!galleryDiv || !statusDiv) {
            console.error("Required DOM elements not found!");
            return;
        }

        console.log("Both DOM elements found. Proceeding...");

        // Fetch wallpapers from Unsplash
        async function fetchWallpapers() {
            try {
                const response = await fetch(
                    `https://api.unsplash.com/photos?per_page=10&client_id=4b2pM-iD5Ltty5GOtVnOs7KDOUopxaTdijfXaDHhXcY`
                );
                const photos = await response.json();
                console.log("Unsplash API Response:", photos);
                return photos;
            } catch (error) {
                console.error("Error fetching wallpapers:", error);
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
                img.src = photo.urls.small;
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
            const response = await fetch(imageUrl);
            const blob = await response.blob();

            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onloadend = () =>
                    resolve(reader.result.split(",")[1]); // Extract Base64
                reader.onerror = reject;
                reader.readAsDataURL(blob);
            });
        }

        // Add event listeners for options
        document
            .getElementById("set-background")
            .addEventListener("click", () => setDocumentBackground(selectedPhoto));
        document
            .getElementById("set-header")
            .addEventListener("click", () => setHeaderImage(selectedPhoto));
        document
            .getElementById("set-footer")
            .addEventListener("click", () => setFooterImage(selectedPhoto));

        // Load wallpapers and render gallery
        const photos = await fetchWallpapers();
        if (photos.length > 0) {
            renderGallery(photos);
        } else {
            statusDiv.textContent = "No photos found. Please try again.";
        }
    });
});
