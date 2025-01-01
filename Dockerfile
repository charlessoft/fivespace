# Use the official Node.js 16 image as a base
FROM node:20

# Set the working directory inside the container
WORKDIR /usr/src/app

# Copy package.json and package-lock.json (if available)


# Install app dependencies
#RUN npm install

# Copy the rest of the application code


# Expose the port on which your app will run
EXPOSE 3000

# Define the command to run your app
CMD ["npm", "start"]