import streamlit as st
import os
from movies_main import (
    load_movie_data,
    clean_movie_data,
    analyze_movies,
    analyze_movies_years,
    export_report,
    generate_charts,
    embed_charts
)

# ---------------- Search Movie ---------------- #
def search_movie(df_movies, movie_name):
    if movie_name.strip() == "":
        return None
    result = df_movies[df_movies['Movie_Name'].str.strip().str.lower() == movie_name.strip().lower()]
    if not result.empty:
        return result.iloc[0]
    return None

# ---------------- Main Streamlit App ---------------- #
def main():
    st.set_page_config(page_title="Movie Data Analyzer", layout="wide")
    st.title("🎬 Movie Ratings & Reviews Analysis System")
    st.markdown("---")

    # Sidebar Navigation
    st.sidebar.title("Navigation")
    page = st.sidebar.selectbox("Choose a page", ["Search Movie", "Data Cleaning", "Generate Analysis Report", "View Charts"])

    # Load movie data
    if 'df_movies' not in st.session_state:
        with st.spinner("Loading movie data..."):
            try:
                df_movies = load_movie_data('movie_data.xlsx')
                df_movies = clean_movie_data(df_movies)
                st.session_state.df_movies = df_movies
                st.success("Movie data loaded successfully!")
            except Exception as e:
                st.error(f"Error loading movie data: {str(e)}")
                return

    df_movies = st.session_state.df_movies

    # ---------------- Search Movie Page ---------------- #
    if page == "Search Movie":
        st.header("🔍 Search Movie Details")
        movie_name = st.text_input("Enter Movie Name:", placeholder="e.g., Animal")
        if st.button("Search"):
            result = search_movie(df_movies, movie_name)
            if result is not None:
                st.success("Movie found!")

    #------------------Display Details-------------------------#
                col1, col2= st.columns(2)
                    
                with col1:
                    st.subheader("Movie Details")
                    if 'Movie_Name' in result.index:
                        st.write(f"**Name:** {result['Movie_Name']}")
                    if 'Directors' in result.index:
                        st.write(f"**Director's Name:** {result['Directors']}")
                    if 'Writers' in result.index:
                        st.write(f"**Writer's Name:** {result['Writers']}")
                    if 'Genre' in result.index:
                        st.write(f"**Genre:** {result['Genre']}")
                    if 'Release_Year' in result.index:
                        st.write(f"**Release Year:** {result['Release_Year']}")
                    if 'overview' in result.index:
                        st.write(f"**Overview:** {result['overview']}")
                with col2:
                    st.subheader("Review details")
                    if 'Reviewer' in result.index:
                        st.write(f"**Reviewer:** {result['Reviewer']}")
                    if 'Ratings' in result.index:
                        st.write(f"**Ratings:** ⭐{result['Ratings']}")
                    if 'Review_Category' in result.index:
                        st.write(f"**Reviw category:** {result['Review_Category']}")
                    if 'Review' in result.index:
                        st.write(f"**Review:** {result['Review']}")


        #---------view details-------------------#
                with st.expander("Show Details"):
                 st.dataframe(result.to_frame().T)
        
        
            else:
                st.error("Movie not found. Please check the name.")

    # ---------------- Data Cleaning & Analysis ---------------- #
    elif page == "Data Cleaning":
       st.header("🧹 Data Cleaning Overview")

       st.markdown("This page shows missing values and cleaning actions applied to the movie dataset.")

       # Show missing values before cleaning
       st.subheader("Missing Values Before Cleaning")
       missing_before = df_movies.isna().sum()
       st.dataframe(missing_before)

       # Show duplicate rows
       st.subheader("Duplicate Rows")
       duplicates = df_movies[df_movies.duplicated()]
       if not duplicates.empty:
           st.dataframe(duplicates)
       else:
           st.write("No duplicate rows found.")
 
       # Show cleaned data preview
           st.subheader("Cleaned Data Preview")
           cleaned_df = clean_movie_data(df_movies)
           st.dataframe(cleaned_df.head(70))

           st.info("Cleaning Steps Applied:\n"
               "- Missing 'Movie_Name', 'Genre', 'Reviewer', 'Reviews', 'Directors', 'Writers' filled with 'Unknown'.\n"
               "- Missing 'Release_Year' and 'Ratings' filled with 0.\n"
               "- Duplicate rows removed.")
        # ---------------- Sorting Section ---------------- #
           st.subheader("🎬 Movies Sorting")

           sort_by = st.selectbox("Sort movies by:", ["Release Year", "Ratings"])
           sort_order = st.radio("Sort Order:", ["Ascending (Low → High)", "Descending (High → Low)"])
 
           if sort_by == "Release Year" and 'Release_Year' in cleaned_df.columns:
            sorted_movies = cleaned_df.sort_values(
                by="Release_Year",
                ascending=True if "Ascending" in sort_order else False
            )
            st.dataframe(sorted_movies[['Release_Year', 'Movie_Name', 'Ratings']].reset_index(drop=True))

           elif sort_by == "Ratings" and 'Ratings' in cleaned_df.columns:
            sorted_movies = cleaned_df.sort_values(
                by="Ratings",
                ascending=True if "Ascending" in sort_order else False
            )
            st.dataframe(sorted_movies[['Ratings', 'Movie_Name', 'Release_Year']].reset_index(drop=True))

           else:
            st.warning(f"⚠️ '{sort_by}' column not found in dataset.")
    
    # ---------------- Generate Analysis Report ---------------- #
    elif page == "Generate Analysis Report":
        st.header("📊 Generate Analysis Report")
        if st.button("Generate Complete Analysis Report"):
            with st.spinner("Generating report..."):
                try:
                    genre_dist, ratings_stats = analyze_movies(df_movies)
                    movie_counts_series, movie_summary = analyze_movies_years(df_movies)
                    
                    output_file = "movies_analysis_report_streamlit.xlsx"
                    export_report(output_file, genre_dist, ratings_stats, movie_counts_series)
                    chart_paths = generate_charts(df_movies)
                    embed_charts(output_file, chart_paths)

                    st.success("Analysis report generated successfully!")
                    
                    # Summary metrics
                    st.subheader("Summary")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Movies", len(df_movies))
                        st.metric("Most Movies released in a year", movie_summary)
                    with col2:
                        st.metric("Average Rating", round(ratings_stats['Average_Rating'],2))
                        st.metric("Max Rating", ratings_stats['Max_Rating'])
                        st.metric("Min Rating", ratings_stats['Min_Rating'])
                    with col3:
                        st.metric("Total Genres", df_movies['Genre'].nunique())
                    
                    # Download link
                    if os.path.exists(output_file):
                        with open(output_file, 'rb') as f:
                            st.download_button(
                                label="📥 Download Analysis Report (Excel)",
                                data=f.read(),
                                file_name=output_file,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                except Exception as e:
                    st.error(f"Error generating report: {str(e)}")

    # ---------------- View Charts ---------------- #
    elif page == "View Charts":
        st.header("📈 Movie Data Visualization Charts")
        charts_dir = 'charts'
        chart_files = {
            'Genre Distribution (Pie Chart)': 'genre_pie.png',
            'Ratings Distribution (Histogram)': 'ratings_hist.png',
            'Movies per Year (Bar Chart)': 'Release_year_movies_bar.png'
        }

        if st.button("Generate Charts"):
            with st.spinner("Generating charts..."):
                try:
                    generate_charts(df_movies, charts_dir)
                    st.success("Charts generated successfully!")
                except Exception as e:
                    st.error(f"Error generating charts: {str(e)}")

        # Display charts
        for chart_name, chart_file in chart_files.items():
            chart_path = os.path.join(charts_dir, chart_file)
            if os.path.exists(chart_path):
                st.subheader(chart_name)
                st.image(chart_path, use_column_width=True)
            else:
                st.info(f"{chart_name} not found. Click 'Generate Charts' to create it.")

if __name__ == "__main__":
    main()
