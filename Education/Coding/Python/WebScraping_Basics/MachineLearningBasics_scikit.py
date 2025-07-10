# -----------------------------------------
# Python_Machine_Learning:
# scikit-learn basics
# -----------------------------------------
#
# Workflow:
# 1. Load data
# 2. Split data
# 3. Train model
# 4. Predict
# 5. Evaluate
#
# Example:
#   from sklearn.datasets import load_iris
#   from sklearn.model_selection import train_test_split
#   from sklearn.tree import DecisionTreeClassifier
#   from sklearn.metrics import accuracy_score
#
#   iris = load_iris()
#   X_train, X_test, y_train, y_test = train_test_split(iris.data, iris.target, test_size=0.3)
#
#   model = DecisionTreeClassifier()
#   model.fit(X_train, y_train)
#   y_pred = model.predict(X_test)
#   print("Accuracy:", accuracy_score(y_test, y_pred))
#
# Common models:
# - Decision Trees, Random Forests
# - Logistic Regression
# - Support Vector Machines
# - K-Nearest Neighbors
# - Clustering (KMeans)
#
# Best practices:
# - Use train/test split.
# - Scale data if needed.
# - Tune hyperparameters.
# - Evaluate with appropriate metrics.
#
# -----------------------------------------
