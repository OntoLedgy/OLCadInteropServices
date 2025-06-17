using CADInteropServices.Objects.AutoCAD.Spaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CADInteropServices.Transformers
{
    public class TransformationMatrix
	{
		public double[,] Matrix { get; private set; }


        public TransformationMatrix()
        {
            // Initialize to identity matrix
            Matrix = new double[,]
            {
                {1, 0, 0, 0 },
                {0, 1, 0, 0 },
                {0, 0, 1, 0 },
                {0, 0, 0, 1 }
            };
        }

        public TransformationMatrix(
			double[] insertionPoint,
			double rotation,
			double[] scaleFactors)
		{
			// Initialize the transformation matrix
			Matrix = CreateTransformationMatrix(
				insertionPoint,
				rotation,
				scaleFactors);

		}

		private double[,] CreateTransformationMatrix(
			double[] insertionPoint,
			double rotation,
			double[] scaleFactors)
		{
			// Create scaling matrix
			double[,] scaleMatrix = {
				{ scaleFactors[0], 0, 0, 0 },
				{ 0, scaleFactors[1], 0, 0 },
				{ 0, 0, scaleFactors[2], 0 },
				{ 0, 0, 0, 1 }
			};

			// Create rotation matrix around Z-axis
			double cosR = Math.Cos(rotation);
			double sinR = Math.Sin(rotation);

			double[,] rotationMatrix = {
				{ cosR, -sinR, 0, 0 },
				{ sinR, cosR,  0, 0 },
				{ 0,    0,     1, 0 },
				{ 0,    0,     0, 1 }
			};

			// Create translation matrix
			double[,] translationMatrix = {
				{ 1, 0, 0, insertionPoint[0] },
				{ 0, 1, 0, insertionPoint[1] },
				{ 0, 0, 1, insertionPoint[2] },
				{ 0, 0, 0, 1 }
			};

			// Combine the transformations: M = T * R * S
			double[,] combinedMatrix = MultiplyMatrices(
				translationMatrix,
				MultiplyMatrices(
					rotationMatrix,
					scaleMatrix)
				);

			return combinedMatrix;
		}

        public Coordinates TransformCoordinates(Coordinates point)
        {
            double[] pointArray = new double[] { point.X, point.Y, point.Z };
            double[] transformedPoint = TransformPoint(pointArray);
            return new Coordinates(transformedPoint[0], transformedPoint[1], transformedPoint[2]);
        }


        public void Translate(double tx, double ty, double tz)
        {
            double[,] translationMatrix = {
                { 1, 0, 0, tx },
                { 0, 1, 0, ty },
                { 0, 0, 1, tz },
                { 0, 0, 0, 1 }
            };

            Matrix = MultiplyMatrices(Matrix, translationMatrix);
        }

        public void RotateZ(double angle)
        {
            double cosR = Math.Cos(angle);
            double sinR = Math.Sin(angle);

            double[,] rotationMatrix = {
                { cosR, -sinR, 0, 0 },
                { sinR, cosR,  0, 0 },
                { 0,    0,     1, 0 },
                { 0,    0,     0, 1 }
            };

            Matrix = MultiplyMatrices(Matrix, rotationMatrix);
        }

        public void Scale(double sx, double sy, double sz)
        {
            double[,] scaleMatrix = {
                { sx, 0,  0, 0 },
                { 0, sy,  0, 0 },
                { 0,  0, sz, 0 },
                { 0,  0,  0, 1 }
            };

            Matrix = MultiplyMatrices(Matrix, scaleMatrix);
        }

   

    public double[] TransformPoint(double[] point)
		{
			double[] result = new double[3];

			double[] pointHomogeneous = { point[0], point[1], point[2], 1 };

			for (int i = 0; i < 3; i++)
			{
				result[i] = Matrix[i, 0] * pointHomogeneous[0] +
							Matrix[i, 1] * pointHomogeneous[1] +
							Matrix[i, 2] * pointHomogeneous[2] +
							Matrix[i, 3] * pointHomogeneous[3];
			}

			return result;
		}

		public TransformationMatrix Combine(TransformationMatrix other)
		{
			double[,] combinedMatrix = MultiplyMatrices(this.Matrix, other.Matrix);
			return new TransformationMatrix(combinedMatrix);
		}

		private TransformationMatrix(double[,] matrix)
		{
			Matrix = matrix;
		}

		private double[,] MultiplyMatrices(double[,] A, double[,] B)
		{
			double[,] result = new double[4, 4];

			for (int i = 0; i < 4; i++)
			{
				for (int j = 0; j < 4; j++)
				{
					result[i, j] = A[i, 0] * B[0, j] +
								   A[i, 1] * B[1, j] +
								   A[i, 2] * B[2, j] +
								   A[i, 3] * B[3, j];
				}
			}

			return result;
		}
	}
}