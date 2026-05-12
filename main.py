import preprocesador as pp
import kmeans

HAM_folder = "KNN-ENTRENAMIENTO/HAM"
resultado_HAM = "PreprocesamientoSteps/HAM"
SPAM_folder = "KNN-ENTRENAMIENTO/SPAM"
resultado_SPAM = "PreprocesamientoSteps/SPAM"

pp.main(HAM_folder, resultado_HAM)
pp.main(SPAM_folder, resultado_SPAM)
pp.join_matrix(resultado_HAM, resultado_SPAM, "PreprocesamientoSteps")

kmeans.main("PreprocesamientoSteps", "KNN-PRUEBA", "PreprocesamientoSteps/Prueba")


