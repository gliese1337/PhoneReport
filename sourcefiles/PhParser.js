/ *   P h P a r s e r . j s  
 * 	 E v e r y t h i n g   f o r   m a n i p u l a t i n g   c a l l   d a t a   a n d   m a k i n g   r e p o r t s  
 * 	 G l o b a l s   U s e d :  
 * 	 	 A c t i v e X O b j e c t ,   C D A T A ,   D a t e ,   N u m b e r ,   p a r s e I n t ,   i s N a n ,   P a r s e I n v o i c e ,  
 * 	 	 E n u m e r a t o r ,   a l e r t ,   t i m e L B o u n d ,   t i m e U B o u n d ,   d u r B o u n d ,   s t a t b o x ,   W S h e l l  
 * 	 G l o b a l s   D e f i n e d :  
 * 	 	 E x p o r t C a l l S t r u c t u r e ,   I m p o r t C a l l S t r u c t u r e ,   I m p o r t D a t a ,   L o a d D a t a ,  
 * 	 	 M a k e D i r e c t o r y ,   g e t C a l l T y p e ,   G e n e r a t e R e p o r t ,   E x p o r t 2 E X L  
 * /  
  
 / *   E x p o r t C a l l S t r u c t u r e ( f s o , c d a t )  
 * 	 A r g u m e n t  
 * 	 	 f s o   ( O b j e c t ) :   a   f i l e s y s t e m   o b j e c t  
 * 	 	 c d a t   ( O b j e c t ) :   a   c a l l   d a t a   s t r u c t u r e   p r o d u c e d   b y   t h e   i n v o i c e   p a r s e r  
 * 	 V a r i a b l e   f s o ,   j ,   o u t f i l e ,   p a t h ,   s p a t h  
 * 	 G l o b a l   N u m b e r  
 * /  
  
 f u n c t i o n   E x p o r t C a l l S t r u c t u r e ( f s o , c d a t ) {  
  
 	 v a r 	 p a t h   =   f s o . G e t A b s o l u t e P a t h N a m e ( " . " ) ,  
 	 	 s p a t h   =   p a t h . s u b s t r i n g ( 0 , p a t h . l a s t I n d e x O f ( " \ \ " ) )   +   " \ \ C a l l   D a t a " ;  
 	 i f ( ! f s o . F o l d e r E x i s t s ( s p a t h ) ) { f s o . C r e a t e F o l d e r ( s p a t h ) ; }  
  
 	 v a r   o u t f i l e   =   f s o . O p e n T e x t F i l e ( s p a t h   +   " \ \ "   +   c d a t . n a m e ,   2 ,   t r u e ,   - 1 ) ;   / / o p e n   f o r   w r i t i n g ,   c r e a t e   n e w   f i l e ,   u n i c o d e   f o r m a t  
  
 	 o u t f i l e . W r i t e L i n e ( " H e a d e r s : " ) ;  
 	 o u t f i l e . W r i t e L i n e ( c d a t . h e a d e r s . j o i n ( ) ) ;  
 	 o u t f i l e . W r i t e L i n e ( " N u m b e r s : " ) ;  
 	 o u t f i l e . W r i t e L i n e ( c d a t . n u m b e r s . j o i n ( ) ) ;  
 	 o u t f i l e . W r i t e L i n e ( " C a l l   D a t a : " ) ;  
 	 f o r ( v a r   j = 0 ; j < c d a t . d a t a . l e n g t h ; j + + ) {  
 	 	 o u t f i l e . W r i t e L i n e ( 	 N u m b e r ( c d a t . d a t a [ j ] [ 0 ] / 1 0 0 0 ) . t o S t r i n g ( 3 6 ) + " \ t " +  
 	 	 	 	 	 N u m b e r ( c d a t . d a t a [ j ] [ 1 ] / 1 0 0 0 + 7 2 0 0 ) . t o S t r i n g ( 3 6 ) + " \ t " +  
 	 	 	 	 	 c d a t . d a t a [ j ] [ 2 ] + " \ t " +   / /   c o u l d   s a v e   a   l i t t l e   m o r e   s p a c e   b y  
 	 	 	 	 	 c d a t . d a t a [ j ] [ 3 ]   / /   n e g a t i n g   t h e   d u r a t i o n   t o   r e m o v e   t h e   " - "   c h a r a c t e r s  
 	 	 	 	 	 ) ;  
 	 }  
 	 o u t f i l e . C l o s e ( ) ;  
 }  
  
 / *   I m p o r t C a l l S t r u c t u r e ( f s o ,   d f n a m e )  
 * 	 A r g u m e n t  
 * 	 	 f s o   ( O b j e c t ) :   a   f i l e   s y s t e m   o b j e c t  
 * 	 	 d f n a m e   ( S t r i n g ) :   t h e   p a t h   o f   a   d a t a   f i l e  
 * 	 V a r i a b l e   D A T A B L O C K ,   d f ,   r e c o r d  
 * 	 G l o b a l   p a r s e I n t  
 * /  
  
 f u n c t i o n   I m p o r t C a l l S t r u c t u r e ( f s o , d f n a m e ) {  
 	 v a r 	 h e a d e r s , n u m b e r s , d a t a = [ ] ,  
 	 	 d f   =   f s o . O p e n T e x t F i l e ( d f n a m e ,   1 ,   f a l s e ,   - 1 ) ;   / / o p e n   f o r   r e a d i n g  
 	 d f . S k i p L i n e ( ) ;  
 	 h e a d e r s   =   d f . R e a d L i n e ( ) . s p l i t ( " , " ) ;  
 	 d f . S k i p L i n e ( ) ;  
 	 n u m b e r s   =   d f . R e a d L i n e ( ) . s p l i t ( " , " ) ;  
 	 d f . S k i p L i n e ( ) ;  
 	 w h i l e ( ! d f . A t E n d O f S t r e a m ) {  
 	 	 v a r   r e c o r d   =   d f . R e a d L i n e ( ) . s p l i t ( " \ t " ) ;  
 	 	 r e c o r d [ 0 ]   =   p a r s e I n t ( r e c o r d [ 0 ] , 3 6 ) * 1 0 0 0 ;  
 	 	 r e c o r d [ 1 ]   =   r e c o r d [ 1 ] = = " N a N " ? " M s g " : ( p a r s e I n t ( r e c o r d [ 1 ] , 3 6 ) - 7 2 0 0 ) * 1 0 0 0 ;  
 	 	 d a t a . p u s h ( r e c o r d ) ;  
 	 }  
 	 d f . C l o s e ( ) ;  
 	 r e t u r n   { " n a m e " : d f n a m e , " h e a d e r s " : h e a d e r s , " n u m b e r s " : n u m b e r s , " d a t a " : d a t a } ;  
 }  
  
 / *   I m p o r t D a t a ( E X L ,   i n v f i l e s )  
 * 	 A r g u m e n t  
 * 	 	 E X L   ( O b j e c t ) :   a n   E x c e l   a p p l i c a t i o n   r e f e r e n c e  
 * 	 	 i n v f i l e s   ( A r r a y   S t r i n g ) :   a   l i s t   o f   p a t h s   t o   i n v o i c e   f i l e s  
 * 	 V a r i a b l e   e x p d a t a ,   i  
 * 	 G l o b a l   C D A T A ,   P a r s e I n v o i c e ,   E x p o r t C a l l S t r u c t u r e  
 * /  
  
 f u n c t i o n   I m p o r t D a t a ( E X L , i n v f i l e s ) {  
 	 i f ( ! i n v f i l e s . l e n g t h ) { r e t u r n ; }  
 	 C D A T A   =   [ ] ;  
 	 v a r   f s o   =   n e w   A c t i v e X O b j e c t ( " S c r i p t i n g . F i l e S y s t e m O b j e c t " ) ;  
 	 f o r ( v a r   i = 0 ; i < i n v f i l e s . l e n g t h ; i + + ) {  
 	 	 v a r   e x p d a t a   =   P a r s e I n v o i c e ( E X L ,   i n v f i l e s [ i ] ) ;  
 	 	 E x p o r t C a l l S t r u c t u r e ( f s o , e x p d a t a ) ;  
 	 	 C D A T A . p u s h ( e x p d a t a ) ;  
 	 }  
 	 f s o   =   n u l l ;  
 }  
  
 / *   L o a d D a t a ( s d a t e ,   e d a t e )  
 * 	 A r g u m e n t  
 * 	 	 s d a t e   ( N u m b e r ) :   t h e   s t a r t i n g   d a t e   f o r   t h e   r e p o r t  
 * 	 	 e d a t e   ( N u m b e r ) :   t h e   e n d i n g   d a t e   f o r   t h e   r e p o r t  
 * 	 V a r i a b l e   c f n a m e ,   c f p a t h ,   c m o n t h ,   c y e a r ,   d a t a f i l e s ,   d a t a f o l d e r ,   e d O b j ,   e m o n t h ,   e y e a r ,   f c ,   f s o ,   f t y p e ,   i ,   j ,   p a t h ,   s d O b j ,   s m o n t h ,   s p a t h ,   s y e a r  
 * 	 G l o b a l   A c t i v e X O b j e c t ,   C D A T A ,   D a t e ,   I m p o r t C a l l S t r u c t u r e ,   p a r s e I n t ,   E n u m e r a t o r  
 * /  
  
 f u n c t i o n   L o a d D a t a ( s d a t e , e d a t e ) {  
 	 v a r 	 i ,  
 	 	 f s o   =   n e w   A c t i v e X O b j e c t ( " S c r i p t i n g . F i l e S y s t e m O b j e c t " ) ,  
 	 	 p a t h   =   f s o . G e t A b s o l u t e P a t h N a m e ( " . " ) ,  
 	 	 s p a t h   =   p a t h . s u b s t r i n g ( 0 , p a t h . l a s t I n d e x O f ( " \ \ " ) )   +   " \ \ C a l l   D a t a " ;  
  
 	 / / G e t   a   l i s t   o f   a l l   o f   t h e   d a t a   f i l e s   c o v e r i n g   t h e   g i v e n   d a t e   r a n g e  
 	 / / T o   A d d -   c h e c k   t o   s e e   i f   t h e r e ' s   e n o u g h   d a t a   t o   c o v e r   t h e   e n t i r e   d a t e   r a n g e  
 	 v a r 	 d a t a f o l d e r   =   f s o . G e t F o l d e r ( s p a t h ) ,  
 	 	 d a t a f i l e s   =   [ ] ,  
 	 	 s d O b j   =   n e w   D a t e ( s d a t e ) ,  
 	 	 e d O b j   =   n e w   D a t e ( e d a t e ) ,  
 	 	 s y e a r   =   s d O b j . g e t F u l l Y e a r ( ) ,  
 	 	 s m o n t h   =   s d O b j . g e t M o n t h ( ) ,  
 	 	 e y e a r   =   e d O b j . g e t F u l l Y e a r ( ) ,  
 	 	 e m o n t h   =   e d O b j . g e t M o n t h ( ) ;  
 	 / / a l e r t ( " S Y e a r :   " + s y e a r + "   S M o n t h :   " + s m o n t h + " \ n E Y e a r :   " + e y e a r + "   E M o n t h :   " + e m o n t h ) ;  
 	 f o r ( v a r   f c   =   n e w   E n u m e r a t o r ( d a t a f o l d e r . f i l e s ) ; ! f c . a t E n d ( ) ; f c . m o v e N e x t ( ) ) {  
 	 	 v a r   c f p a t h   =   f c . i t e m ( ) + " " ;  
 	 	 v a r   c f n a m e   =   c f p a t h . s u b s t r ( c f p a t h . l a s t I n d e x O f ( " \ \ " ) + 1 ) ;  
 	 	 v a r   f t y p e   =   c f n a m e . s p l i t ( " . " ) ;  
 	 	 i f ( f t y p e . l e n g t h < 2   | |   f t y p e [ 0 ] . l e n g t h < 6 ) { c o n t i n u e ; }  
 	 	 i f ( f t y p e [ 1 ] = = " c d a t " ) {  
 	 	 	 v a r   c y e a r   =   p a r s e I n t ( f t y p e [ 0 ] . s u b s t r ( 0 , 4 ) , 1 0 ) ;  
 	 	 	 v a r   c m o n t h   =   p a r s e I n t ( f t y p e [ 0 ] . s u b s t r ( 4 ) , 1 0 ) - 1 ;  
 	 	 	 / / a l e r t ( " Y e a r :   " + c y e a r + "   M o n t h :   " + c m o n t h ) ;  
 	 	 	 / / a l e r t ( ( c y e a r   > =   s y e a r )   & &   ( c y e a r   < =   e y e a r )   & &   ( c m o n t h   > =   s m o n t h )   & &   ( c m o n t h   < =   e m o n t h ) ) ;  
 	 	 	 i f ( ( c y e a r   > =   s y e a r )   & &   ( c y e a r   < =   e y e a r )   & &   ( c m o n t h   > =   s m o n t h )   & &   ( c m o n t h   < =   e m o n t h ) ) { d a t a f i l e s . p u s h ( c f p a t h ) ; }  
 	 	 }  
 	 }  
 	 i f ( ! d a t a f i l e s . l e n g t h ) {  
 	 	 a l e r t ( " N o   d a t a   f i l e s   f o u n d   f o r   t h e   s p e c i f i e d   t i m e   r a n g e . " ) ;  
 	 	 r e t u r n   f a l s e ;  
 	 }  
  
 	 / / f o r   e a c h   d a t a   f i l e ,   c h e c k   t o   s e e   i f   t h e r e ' s   a   r e c o r d   w i t h   t h e   s a m e   n a m e   a l r e a d y   l o a d e d ;  
 	 / / i f   n o t ,   o p e n   i t   a n d   r e a d   i n   t h e   d a t a  
 	 / / C D A T A   =   [ { " n a m e " : " " , " h e a d e r s " : n e w   A r r a y ( ) , " n u m b e r s " : n e w   A r r a y ( ) , " d a t a " : n e w   A r r a y ( ) } ] ;  
 	 i f ( ! C D A T A ) { C D A T A   =   [ ] ; }  
 	 f o r ( i = 0 ; i < d a t a f i l e s . l e n g t h ; i + + ) {  
 	 	 v a r   j ;  
 	 	 f o r ( j = 0 ; j < C D A T A . l e n g t h ; j + + ) { i f ( d a t a f i l e s [ i ]   = =   C D A T A [ j ] . n a m e ) { b r e a k ; } }  
 	 	 i f ( j = = C D A T A . l e n g t h ) { C D A T A . p u s h ( I m p o r t C a l l S t r u c t u r e ( f s o , d a t a f i l e s [ i ] ) ) ; }  
 	 }  
 	 C D A T A . s o r t ( f u n c t i o n ( a , b ) { r e t u r n   p a r s e I n t ( a . n a m e , 1 0 )   <   p a r s e I n t ( b . n a m e , 1 0 ) ; } ) ;  
 	 f s o   =   n u l l ;  
 	 r e t u r n   t r u e ;  
 }  
  
 / *   M a k e D i r e c t o r y ( E X L ,   D i r S h e e t P a t h )  
 * 	 A r g u m e n t  
 * 	 	 E X L   ( O b j e c t ) :   a n   E x c e l   a p p l i c a t i o n   r e f e r e n c e  
 * 	 	 D i r S h e e t P a t h   ( S t r i n g ) :   P a t h   t o   a   d i r e c t o r y   f i l e  
 * 	 V a r i a b l e   a r e a ,   d i r ,   d n a m e ,   d s h e e t ,   i n c ,   l s h e e t ,   m 1 ,   m 2 ,   p h n u m  
 *  
 * 	 G e n e r a t e s   h a s h e s   c o n t a i n i n g   p h o n e   d i r e c t o r y   i n f o r m a t i o n .  
 *  
 * 	 L o o p   t h r o u g h   t h e   d i r e c t o r y   s h e e t   u n t i l   t h e r e   i s   n o   m o r e   d a t a ,  
 * 	 g e n e r a t i n g   a   n a m e   t o   a s s o c i a t e   w i t h   e a c h   n u m b e r   b a s e d   o n   t h e    
 * 	 a r e a   n a m e   a n d   c o m p a n i o n   n a m e s .   S k i p   a n y   p h o n e   n u m b e r s   m a r k e d  
 * 	 a s   " E x t r a " .   A l s o   a s s o c i a t e   c a l l i n g - t r e e   c o d e s   w i t h   n u m b e r s .  
 *  
 * 	 R e t u r n s   a   2 - e l e m e n t   a r r a y   c o n t a i n i n g   h a s h e s   a s s o c i a t i n g   p h o n e  
 * 	 n u m b e r s   w i t h   n a m e s   a n d   c a l l i n g - t r e e   c o d e s .  
 * /  
  
 f u n c t i o n   M a k e D i r e c t o r y ( E X L ,   D i r S h e e t P a t h ) {  
 	 v a r 	 i n c = 4 ,  
 	 	 d i r   =   [ [ ] , [ ] ] ,  
 	 	 l s h e e t   =   E X L . W o r k b o o k s . O p e n ( D i r S h e e t P a t h ) ,  
 	 	 d s h e e t   =   l s h e e t . W o r k S h e e t s ( 1 / * " C L i s t " * / ) ,  
 	 	 a r e a ,  
 	 	 p h n u m ,  
 	 	 d n a m e ,  
 	 	 m 1 , m 2 ;  
  
 	 w h i l e ( d s h e e t . C e l l s ( i n c ,   1 ) . T e x t ! = = " " ) {  
 	 	 a r e a = d s h e e t . C e l l s ( i n c , 2 ) . T e x t ;  
 	 	 i f ( a r e a . t o L o w e r C a s e ( ) ! = " e x t r a "   & &   a r e a ! = = " " ) {  
 	 	 	 m 1 = d s h e e t . C e l l s ( i n c , 3 ) . T e x t ;  
 	 	 	 m 2 = d s h e e t . C e l l s ( i n c , 4 ) . T e x t . s u b s t r ( 0 , 1 2 ) ;  
 	 	 	 i f ( m 1 . t o L o w e r C a s e ( ) = = " c l o s e d " ) { m 1 = " " ; }  
 	 	 	 i f ( m 2 = = = " " ) { d n a m e = ( m 1 = = = " " ) ? a r e a . s u b s t r ( 0 , 3 1 ) : m 1 . s u b s t r ( 0 , 3 0 ) ; }  
 	 	 	 e l s e {  
 	 	 	 	 m 1 = m 1 . s u b s t r ( 0 , 1 2 ) ;  
 	 	 	 	 d n a m e   =   a r e a . s u b s t r ( 0 , 3 1 - ( m 1 . l e n g t h + m 2 . l e n g t h + 2 ) ) + "   " + m 1 + "   " + m 2 ;  
 	 	 	 }  
 	 	 	 p h n u m   =   d s h e e t . C e l l s ( i n c , 1 ) . T e x t ;  
 	 	 	 p h n u m   =   p h n u m . s u b s t r i n g ( p h n u m . l e n g t h - 9 ) ;  
 	 	 	 d i r [ 0 ] [ p h n u m ]   =   d n a m e ;  
 	 	 	 d i r [ 1 ] [ p h n u m ]   =   d s h e e t . C e l l s ( i n c , 5 ) . T e x t + " " ;   / / A l l o w e d   c a l l s   c o d e ,   f o r c e   s t r i n g   j u s t   i n   c a s e  
 	 	 }  
 	 	 i n c + + ;  
 	 }  
  
 	 l s h e e t . C l o s e ( f a l s e ) ;  
 	 r e t u r n   d i r ;  
 }  
  
 / *   g e t C a l l T y p e ( c o d e 1 ,   c o d e 2 ,   t i m e ,   d u r a t i o n )  
 * 	 A r g u m e n t  
 * 	 	 c o d e 1 , c o d e 2   ( S t r i n g ) : 	 T h e   c a l l i n g   c o d e s   f o r   t h e   n u m b e r s   i n v o l v e d   i n   t h e   c a l l  
 * 	 	 t i m e   ( N u m b e r ) : 	 	 T h e   t i m e   w h e n   t h e   c a l l   w a s   m a d e  
 * 	 	 d u r a t i o n   ( N u m b e r ) : 	 T h e   l e n g t h   o f   t h e   c a l l  
 * 	 V a r i a b l e   c t ,   i c ,   o c ,   t e m p  
 * 	 G l o b a l   i s N a N ,   t i m e L B o u n d ,   t i m e U B o u n d ,   d u r B o u n d  
 *  
 * 	 D e t e r m i n e s   t h e   a l l o w e d n e s s   o f   p h o n e   c a l l s   v i a   p o s i t i o n s   i n   t h e   c a l l i n g   t r e e .  
 *  
 * 	 R e t u r n s   a   s t r i n g   c o n t a i n i n g   a   n u m b e r   o r   " N "   ( p l u s   t h e   o p t i o n a l   c o l o r   c o d e )  
 * 	 i n d i c a t i n g   t h e   c l a s s i f i c a t i o n   o f   t h e   c a l l   b a s e d   o n   t i m e   a n d   r e l a t i v e   c a l l i n g -  
 * 	 t r e e   p o s i t i o n s .  
 *  
 * 	 N .   S p e c i a l l y   t r a c k e d   n u m b e r  
 * 	 0 .   C a l l   b e t w e e n   m i s s i o n   a n d   n o n - m i s s i o n  
 * 	 1 .   C a l l   b e t w e e n   z o n e   l e a d e r s   a n d   z o n e   m e m b e r s  
 * 	 2 .   C a l l   b e t w e e n   d i s t r i c t   l e a d e r s   a n d   d i s t r i c t   m e m b e r s  
 * 	 3 .   C a l l   b e t w e e n   d i s t r i c t   m e m b e r s  
 * 	 4 .   C a l l   b e t w e e n   s i s t e r s  
 * 	 5 .   C a l l   b e t w e e n   e l d e r s  
 * 	 6 .   C a l l   b e t w e e n   e l d e r s   &   s i s t e r s  
 * 	 7 .   U n i v e r s a l l y   a l l o w e d   c a l l  
 *  
 * 	 R e t u r n   V a l u e s :  
 *  
 * 	 0 - 3 :   t i m e   i s s u e s   ( 0 0 0 0 0 0 0 0   -   1 1 0 0 0 0 0 0 )  
 * 	 r e m a i n i n g   d i g i t s :   m i s s i o n   c a l l   t r e e  
 *  
 * 	 0   a l l   g o o d 	 1 6   I n t e r s i s t e r  
 * 	 1   o u t   o f   t i m e 	 1 7   ( 1 6   1 )  
 * 	 2   o v e r   t i m e 	 1 8   ( 1 6   2 )  
 * 	 3   ( 2   1 ) 	 	 1 9   ( 1 6   2   1 )  
 * 	 4   Z L - Z m 	 	 2 0   I n t e r e l d e r  
 * 	 5   ( 4   1 ) 	 	 2 1   ( 1 6   4   1 )  
 * 	 6   ( 4   2 ) 	 	 2 2   ( 1 6   4   2 )  
 * 	 7   ( 4   2   1 ) 	 2 3   ( 1 6   4   2   1 )  
 * 	 8   D L - D m 	 	 2 4   E l d e r   t o   S i s t e r  
 * 	 9   ( 8   1 ) 	 	 2 5   ( 1 6   8   1 )  
 * 	 1 0   ( 8   2 ) 	 2 6   ( 1 6   8   2 )  
 * 	 1 1   ( 8   2   1 ) 	 2 7   ( 1 6   8   2   1 )  
 * 	 1 2   I n   d i s t r i c t 	 2 8   U n i v e r s a l   P r i v e l e g e s  
 * 	 1 3   ( 8   4   1 ) 	 2 9   ( 1 6   8   4   1 )  
 * 	 1 4   ( 8   4   2 ) 	 3 0   ( 1 6   8   4   2 )  
 * 	 1 5   ( 8   4   2   1 ) 	 3 1   ( 1 6   8   4   2   1 )  
 * /  
  
 f u n c t i o n   g e t C a l l T y p e ( c o d e 1 , c o d e 2 , t i m e , d u r a t i o n ) {  
 	 i f ( c o d e 1 . s u b s t r ( 0 , 1 )   = =   " N " ) { r e t u r n   c o d e 1 ; }  
 	 i f ( c o d e 2 . s u b s t r ( 0 , 1 )   = =   " N " ) { r e t u r n   c o d e 2 ; }  
 	 v a r   c t   =   ( t i m e < t i m e L B o u n d | | t i m e > t i m e U B o u n d ) ? 1 : 0 ;  
 	 i f ( ! i s N a N ( d u r a t i o n )   & &   d u r a t i o n > d u r B o u n d ) { c t | = 2 ; }  
 	 i f ( c o d e 1   = =   " F "   | |   c o d e 2   = =   " F " ) { r e t u r n   c t + " " ; }  
 	 i f ( c o d e 1   = =   " U "   | |   c o d e 2   = =   " U "   | |   c o d e 1   = =   " T "   | |   c o d e 2   = = " T " ) { r e t u r n   ( c t | 2 8 ) + " " ; }   / / u n i v e r s a l   c a l l   t o / f r o m   A P s / O f f i c e n i k s / S e n i o r s  
 	 i f ( c o d e 1   = =   c o d e 2 ) { r e t u r n   ( c t | 1 2 ) + " " ; }   / / i n t r a d i s t r i c t   c a l l  
  
 	 v a r   i c   =   c o d e 1 . s p l i t ( " . " ) ;  
 	 v a r   o c   =   c o d e 2 . s p l i t ( " . " ) ;  
 	 i f ( o c . l e n g t h > i c . l e n g t h ) {   / / m a k e   s u r e   o c   i s   s h o r t e r   o r   e q u a l  
 	 	 v a r   t e m p   =   o c ;  
 	 	 o c   =   i c ;  
 	 	 i c   =   t e m p ;  
 	 }  
 	 s w i t c h ( o c . l e n g t h ) {  
 	 	 c a s e   1 :   s w i t c h ( i c . l e n g t h ) {   / / Z o n e   L e a d e r s  
 	 	 	 	 c a s e   2 : 	 i f ( i c [ 0 ] = = o c [ 0 ] ) { r e t u r n   ( c t | 4 ) + " " ; }   / / Z L   c a l l i n g   a   D L   i n   h i s   z o n e  
 / * c a s e   2   e l s e -   f a l l   t h r o u g h * / 	 c a s e   1 : 	 r e t u r n   ( c t | 2 0 ) + " " ;   / / c a l l i n g   a n o t h e r   Z o n e   L e a d e r ,   i n t e r e l d e r  
 	 	 	 	 c a s e   3 :   r e t u r n   ( c t | ( 	 ( i c [ 0 ] = = o c [ 0 ] ) ? 4 :   / / Z L   c a l l i n g   w i t h i n   h i s   z o n e  
 	 	 	 	 	 	 	 ( i c [ 2 ] = = " S " ) ? 2 4 :   / / Z L   c a l l i n g   a   s i s t e r   n o t   i n   h i s   z o n e  
 	 	 	 	 	 	 	 	 2 0 )   / / i n t e r e l d e r  
 	 	 	 	 	 	 ) + " " ;  
 	 	 	 }  
 	 	 c a s e   2 :   s w i t c h ( i c . l e n g t h ) {   / / D i s t r i c t   L e a d e r s  
 	 	 	 	 c a s e   3 :   i f ( i c [ 0 ] = = o c [ 0 ]   & &   i c [ 1 ] = = o c [ 1 ] ) { r e t u r n   ( c t | 8 ) + " " ; }   / / D L   c a l l i n g   w i t h i n   h i s   d i s t r i c t  
 / * c a s e   3   e l s e -   f a l l   t h r o u g h * / 	 	 i f ( i c [ 2 ] = = " S " ) { r e t u r n   ( c t | 2 4 ) + " " ; }   / / D L   c a l l i n g   a   s i s t e r   n o t   i n   h i s   d i s t r i c t  
 	 	 	 	 c a s e   2 :   r e t u r n   ( c t | 2 0 ) + " " ; 	 / / c a l l i n g   a n o t h e r   D i s t r i c t   L e a d e r ,   i n t e r e l d e r  
 	 	 	 }  
 	 	 c a s e   3 :   r e t u r n   ( c t | ( 	 ( i c [ 2 ] = = " S "   & &   o c [ 2 ] = = " S " ) ? 1 6 :   / / i n t e r s i s t e r  
 	 	 	 	 	 ( i c [ 2 ] = = " E "   & &   o c [ 2 ] = = " E " ) ? 2 0 :   / / i n t e r e l d e r  
 	 	 	 	 	 	 	 	 2 4 )   / / e l d e r   t o   s i s t e r  
 	 	 	 	 ) + " " ;  
 	 }  
 }  
  
  
  
  
 / *   G e n e r a t e R e p o r t ( E X L ,   d i r s ,   d a t e s ,   n l i s t ,   c d a t s ,   r t y p e )  
 * 	 A r g u m e n t  
 * 	 	 E X L   ( O b j e c t ) :   a n   E x c e l   a p p l i c a t i o n   o b j e c t  
 * 	 	 d i r s   ( A r r a y ) :   a   l i s t   o f   d i r e c t o r i e s  
 * 	 	 d a t e s   ( A r r a y ) :   a   l i s t   o f   d a t e s   m a r k i n g   o f f   t h e   p e r i o d s   o f   u s e   f o r   e a c h   d i r e c t o r y  
 * 	 	 n l i s t   ( A r r a y ) :   a   l i s t   o f   n u m b e r s   t o   m a k e   r e p o r t s   f o r ;   i f   e m p t y ,   t h e   d i r e c t o r i e s   w i l l   b e   u s e d   t o   m a k e   t h e   l i s t   a u t o m a t i c a l l y  
 * 	 	 c d a t s   ( A r r a y ) :   a   l i s t   o f   c a l l   d a t a   s t r u c t u r e s  
 * 	 	 r t y e p   ( I n t ) :   w h a t   k i n d   o f   r e p o r t   t o   m a k e -     f u l l ,   s u m m a r y ,   o r   b o t h  
 * 	 C l o s u r e   N U M B E R S ,   R E P O R T ,   c d i r ,   d a t i d x ,   i ,   k e y l e n ,   o u t f i l e  
 * 	 V a r i a b l e   D i r R e p ,   c a l l t y p e ,   c u r r e c ,   d u r a t i o n ,   f s o ,   i c o d e ,   i k e y ,   i n n u m ,  
 	 	 j ,   k e y ,   o c o d e ,   o k e y ,   o u t n u m ,   p a t h ,   t i m e ,   u p d a t e R e p o r t B l o c k s  
 * 	 G l o b a l   A c t i v e X O b j e c t ,   s t a t b o x ,   M a k e D i r e c t o r y ,   g e t C a l l T y p e  
 * /  
  
 f u n c t i o n   G e n e r a t e R e p o r t ( E X L , d i r s , d a t e s , n l i s t , c d a t s , r t y p e ) {  
 	 / * d e b u g g i n g   s t u f f * /  
 	 v a r   f s o   =   n e w   A c t i v e X O b j e c t ( " S c r i p t i n g . F i l e S y s t e m O b j e c t " ) ;  
 	 v a r   p a t h   =   f s o . G e t A b s o l u t e P a t h N a m e ( " . " ) ;  
 	 v a r   o u t f i l e   =   f s o . O p e n T e x t F i l e ( p a t h   +   " \ \ l o g f i l e . t x t " ,   2 ,   t r u e ,   - 1 ) ;   / / o p e n   f o r   w r i t i n g ,   c r e a t e   n e w   f i l e ,   u n i c o d e   f o r m a t  
 	 / * d e b u g g i n g   s t u f f * /  
  
 	 v a r 	 i , j , c d i r ,  
 	 	 d a t i d x   =   0 ,  
 	 	 D i r R e p   =   ( n l i s t . l e n g t h   = = =   0 ) ,  
 	 	 k e y l e n ,  
 	 	 c u r r e c ,  
 	 	 t i m e , d u r a t i o n , c d u r ,  
 	 	 o u t n u m , i n n u m ,  
 	 	 c a l l t y p e ,  
 	 	 i k e y , o k e y ,  
 	 	 i c o d e , o c o d e ,  
 	 	 p h n a m e ,  
 	 	 p a r s e L i n e I n ,  
 	 	 p a r s e L i n e O u t ,  
 	 	 R E P O R T   =   { } , / / R E P O R T   o b j e c t   s t r u c t u r e :   " n a m e " :   [ [ d a t e , t i m e , d u r a t i o n , i n / o u t , p h o n e   n u m b e r , c a l l   t y p e ] ]  
 	 	 N U M B E R S   =   { } ;  
  
 	 f u n c t i o n   a d d S u m I n ( ) {  
 	 	 i f ( N U M B E R S [ i k e y ] . s u m m a r y [ p h n a m e ] ) {  
 	 	 	 N U M B E R S [ i k e y ] . s u m m a r y [ p h n a m e ] [ 0 ] + + ;  
 	 	 	 N U M B E R S [ i k e y ] . s u m m a r y [ p h n a m e ] [ 1 ] + = c d u r + 7 2 0 0 0 0 0 ;  
 	 	 } e l s e { / / i n c o u n t , i n t i m e , o u t c o u n t , o u t t i m e  
 	 	 	 N U M B E R S [ i k e y ] . s u m m a r y [ p h n a m e ]   =   [ 1 , c d u r , 0 , - 7 2 0 0 0 0 0 ] ;  
 	 	 } 	 	 	 	    
 	 }  
 	 f u n c t i o n   a d d S u m O u t ( ) {  
 	 	 i f ( N U M B E R S [ o k e y ] . s u m m a r y [ p h n a m e ] ) {  
 	 	 	 N U M B E R S [ o k e y ] . s u m m a r y [ p h n a m e ] [ 2 ] + + ;  
 	 	 	 N U M B E R S [ o k e y ] . s u m m a r y [ p h n a m e ] [ 3 ] + = c d u r + 7 2 0 0 0 0 0 ;  
 	 	 } e l s e {  
 	 	 	 N U M B E R S [ o k e y ] . s u m m a r y [ p h n a m e ]   =   [ 0 , - 7 2 0 0 0 0 0 , 1 , c d u r ] ;  
 	 	 }  
 	 }  
 	 f u n c t i o n   u p d a t e C a l l T y p e I n ( ) {  
 	 	 i c o d e   =   ( i k e y   i n   c d i r [ 1 ] ) ? c d i r [ 1 ] [ i k e y ] : " F " ;  
 	 	 o c o d e   =   " F " ;  
 	 	 p h n a m e   =   o u t n u m ;  
 	 	 f o r ( o k e y = o u t n u m ; o k e y . l e n g t h & & o k e y . l e n g t h > = k e y l e n ; o k e y = o k e y . s u b s t r ( 1 ) ) { i f ( o k e y   i n   c d i r [ 1 ] ) {  
 	 	 	 o c o d e   =   c d i r [ 1 ] [ o k e y ] ;  
 	 	 	 p h n a m e   =   c d i r [ 0 ] [ o k e y ] ;  
 	 	 	 b r e a k ;  
 	 	 } }  
 	 	 c a l l t y p e   =   g e t C a l l T y p e ( i c o d e , o c o d e , t i m e , d u r a t i o n ) ; 	 	 	 	    
 	 }  
 	 f u n c t i o n   u p d a t e C a l l T y p e O u t ( ) {  
 	 	 o c o d e   =   ( o k e y   i n   c d i r [ 1 ] ) ? c d i r [ 1 ] [ o k e y ] : " F " ;  
 	 	 i c o d e   =   " F " ;  
 	 	 p h n a m e   =   i n n u m ;  
 	 	 f o r ( i k e y = i n n u m ; i k e y . l e n g t h & & i k e y . l e n g t h > = k e y l e n ; i k e y = i k e y . s u b s t r ( 1 ) ) { i f ( i k e y   i n   c d i r [ 1 ] ) {  
 	 	 	 i c o d e   =   c d i r [ 1 ] [ i k e y ] ;  
 	 	 	 p h n a m e   =   c d i r [ 0 ] [ i k e y ] ;  
 	 	 	 b r e a k ;  
 	 	 } }  
 	 	 c a l l t y p e   =   g e t C a l l T y p e ( i c o d e , o c o d e , t i m e , d u r a t i o n ) ;  
 	 }  
  
 	 / *   " u p d a t e R e p o r t B l o c k s " ( )  
 	 * 	 V a r i a b l e   e l t ,   e n t r y ,   e r r s t r i n g ,   k ,   l ,   p h o n e n a m e  
 	 * 	 O u t e r   N U M B E R S ,   R E P O R T ,   c d a t s ,   c d i r ,   d a t i d x ,   i ,   k e y l e n ,   n l i s t ,   o u t f i l e  
 	 *  
 	 * 	 M a k e / U p d a t e   t h e   r e p o r t   s t r u c t u r e   b a s e d   o n   o u r   r e p o r t i n g   l i s t   &   t h e   c u r r e n t   d a t a   b l o c k  
 	 * 	 R u n s   e v e r y   t i m e   w e   c h a n g e   d a t a b l o c k s   o r   d i r e c t o r i e s  
 	 * /  
 	 f u n c t i o n   u p d a t e R e p o r t B l o c k s ( ) {  
 	 	 v a r   e r r s t r i n g , z ;  
 	 	 i = 0 ;  
 	 	 k e y l e n = 0 ;  
 	 	 f o r ( v a r   l = 0 , z = n l i s t . l e n g t h ; l < z ; l + + ) {  
 	 	 	 v a r   k , w , e n t r y , e l t = n l i s t [ l ] ;  
 	 	 	 / / a l e r t ( " U p d a t i n g   r e p o r t   b l o c k s :   " + e l t ) ;  
 	 	 	 f o r   ( k = 0 , w = c d a t s [ d a t i d x ] . h e a d e r s . l e n g t h ;   k   <   w   ;   k + + ) {   / / c h e c k   i f   i t ' s   i n   t h e   h e a d e r s  
 	 	 	 	 e n t r y   =   c d a t s [ d a t i d x ] . h e a d e r s [ k ] ;  
 	 	 	 	 i f ( e n t r y . l e n g t h   >   e l t . l e n g t h ) { e n t r y   =   e n t r y . s u b s t r ( e n t r y . l e n g t h - e l t . l e n g t h ) ; }  
 	 	 	 	 i f ( e n t r y   = =   e l t ) { b r e a k ; }  
 	 	 	 } i f ( k = = w ) {  
 	 	 	 	 e r r s t r i n g   =   e l t + "   n o t   f o u n d   i n   d a t a   h e a d e r s   f o r   m o n t h   " + ( d a t i d x + 1 ) ;  
 	 	 	 	 s t a t b o x . v a l u e + = e r r s t r i n g + " \ n " ;  
 	 	 	 	 o u t f i l e . W r i t e L i n e ( e r r s t r i n g ) ;  
 	 	 	 }  
  
 	 	 	 / / c h e c k   t o   s e e   i f   e a c h   n u m b e r   a c t u a l l y   a p p e a r s   i n   t h i s   b l o c k   o f   d a t a  
 	 	 	 f o r   ( k = 0 , w = c d a t s [ d a t i d x ] . n u m b e r s . l e n g t h ;   k   <   w ;   k + + ) {  
 	 	 	 	 e n t r y   =   c d a t s [ d a t i d x ] . n u m b e r s [ k ] ;  
 	 	 	 	 i f (  
 	 	 	 	 	 ( ( e n t r y . l e n g t h   >   e l t . l e n g t h )   & &   ( e l t   = =   e n t r y . s u b s t r ( e n t r y . l e n g t h - e l t . l e n g t h ) ) )  
 	 	 	 	 	 | |   ( ( e l t . l e n g t h   >   e n t r y . l e n g t h )   & &   ( e n t r y   = =   e l t . s u b s t r ( e l t . l e n g t h - e n t r y . l e n g t h ) ) )  
 	 	 	 	 	 | |   ( e n t r y   = =   e l t )  
 	 	 	 	 ) { b r e a k ; }  
 	 	 	 } i f ( k ! = w ) { / / o n l y   i f   a   n u m b e r   a p p e a r s   i n   t h e   d a t a ,   m a k e   a   d a t a b l o c k   f o r   i t  
 	 	 	 / *  
 	 	 	 	 R E P O R T   c o n t a i n s   d a t a b l o c k s   i d e n t i f i e d   b y   p h o n e   n a m e s   f o r   o u t p u t t i n g ,   b u t   N U M B E R S   c o n t a i n s   r e f e r e n c e s  
 	 	 	 	 i n d e x e d   b y   p h o n e   n u m b e r   t o   t h e   s a m e   b l o c k s .   E v e r y   t i m e   t h r o u g h   t h e   d i r e c t o r y   l o o p ,   c r e a t e   n e w   d a t a    
 	 	 	 	 b l o c k s   f o r   a n y   p h o n e   n a m e s   t h a t   d o n ' t   e x i s t   y e t ,   a n d   r e - a s s i g n   t h e   n u m b e r   k e y s   t o   p o i n t   t o   t h e  
 	 	 	 	 a p p r o p r i a t e   b l o c k s   a c c o r d i n g   t o   t h e   n a m e   t h a t   p h o n e   n u m b e r   h a s   f o r   t h a t   t r a n s f e r   p e r i o d  
 	 	 	 * /  
 	 	 	 	 v a r   p h o n e n a m e   =   ( e l t   i n   c d i r [ 0 ] ) ? c d i r [ 0 ] [ e l t ] : e l t ;  
 	 	 	 	 i f ( ! ( p h o n e n a m e   i n   R E P O R T ) ) {  
 	 	 	 	 	 o u t f i l e . W r i t e L i n e ( " M a k i n g   d a t a   b l o c k   f o r   "   +   p h o n e n a m e   +   " :   "   +   e l t ) ;  
 	 	 	 	 	 R E P O R T [ p h o n e n a m e ]   =   { " l i s t " : [ ] , " s u m m a r y " : [ ] } ;  
 	 	 	 	 }  
 	 	 	 	 N U M B E R S [ e l t ]   =   R E P O R T [ p h o n e n a m e ] ;  
 	 	 	 	 i f ( ! k e y l e n   | |   k e y l e n   >   e l t . l e n g t h ) { k e y l e n   =   e l t . l e n g t h ; }  
 	 	 	 } e l s e {  
 	 	 	 	 e r r s t r i n g   =   e l t + "   m i s s i n g   f r o m   d a t a   f o r   m o n t h   " + ( d a t i d x + 1 ) ;  
 	 	 	 	 s t a t b o x . v a l u e + = e r r s t r i n g + " \ n " ;  
 	 	 	 	 o u t f i l e . W r i t e L i n e ( e r r s t r i n g ) ;  
 	 	 	 }  
 	 	 }  
 	 }  
  
 	 s w i t c h ( r t y p e & 3 ) {  
 	 	 c a s e   1 :   p a r s e L i n e I n   =   f u n c t i o n ( ) {   / / f u l l   r e p o r t   o n l y  
 	 	 	 	 u p d a t e C a l l T y p e I n ( ) ;  
 	 	 	 	 N U M B E R S [ i k e y ] . l i s t . p u s h ( [ c u r r e c [ 0 ] , d u r a t i o n , " I n c o m i n g " , ( o k e y   i n   c d i r [ 0 ] ? c d i r [ 0 ] [ o k e y ] : o u t n u m ) , c a l l t y p e ] ) ;    
 	 	 	 } ;  
 	 	 	 p a r s e L i n e O u t   =   f u n c t i o n ( ) {  
 	 	 	 	 i f ( c a l l t y p e   = =   " X " ) { u p d a t e C a l l T y p e O u t ( ) ; }  
 	 	 	 	 N U M B E R S [ o k e y ] . l i s t . p u s h ( [ c u r r e c [ 0 ] , d u r a t i o n , " O u t g o i n g " , ( i k e y   i n   c d i r [ 0 ] ? c d i r [ 0 ] [ i k e y ] : i n n u m ) , c a l l t y p e ] ) ;  
 	 	 	 } ;  
 	 	 b r e a k ;  
 	 	 c a s e   2 :   p a r s e L i n e I n   =   f u n c t i o n ( ) {   / / s u m m a r y   r e p o r t   o n l y  
 	 	 	 	 f o r ( o k e y = o u t n u m ; o k e y . l e n g t h & & o k e y . l e n g t h > = k e y l e n ; o k e y = o k e y . s u b s t r ( 1 ) ) { i f ( o k e y   i n   c d i r [ 1 ] ) { b r e a k ; } }  
 	 	 	 	 p h n a m e   =   c d i r [ 0 ] [ o k e y ]   | |   o u t n u m ;  
 	 	 	 	 a d d S u m I n ( ) ;  
 	 	 	 } ;  
 	 	 	 p a r s e L i n e O u t   =   f u n c t i o n ( ) {  
 	 	 	 	 f o r ( i k e y = i n n u m ; i k e y . l e n g t h & & i k e y . l e n g t h > = k e y l e n ; i k e y = i k e y . s u b s t r ( 1 ) ) { i f ( i k e y   i n   c d i r [ 1 ] ) { b r e a k ; } }  
 	 	 	 	 p h n a m e   =   c d i r [ 0 ] [ i k e y ]   | |   i n n u m ;  
 	 	 	 	 a d d S u m O u t ( ) ;  
 	 	 	 } ;  
 	 	 b r e a k ;  
 	 	 c a s e   3 :   p a r s e L i n e I n   =   f u n c t i o n ( ) {   / / b o t h  
 	 	 	 	 u p d a t e C a l l T y p e I n ( ) ;  
 	 	 	 	 a d d S u m I n ( ) ;  
 	 	 	 	 N U M B E R S [ i k e y ] . l i s t . p u s h ( [ c u r r e c [ 0 ] , d u r a t i o n , " I n c o m i n g " , p h n a m e , c a l l t y p e ] ) ;    
 	 	 	 	 / / a s s i g n   b y   r e f e r e n c e   m e a n s   t h a t   u p d a t i n g   N U M B E R S   a u t o m a t i c a l l y   u p d a t e s   R E P O R T  
 	 	 	 } ;  
 	 	 	 p a r s e L i n e O u t   =   f u n c t i o n ( ) {  
 	 	 	 	 i f ( c a l l t y p e   = =   " X " ) { u p d a t e C a l l T y p e O u t ( ) ; }  
 	 	 	 	 a d d S u m O u t ( ) ;  
 	 	 	 	 N U M B E R S [ o k e y ] . l i s t . p u s h ( [ c u r r e c [ 0 ] , d u r a t i o n , " O u t g o i n g " , p h n a m e , c a l l t y p e ] ) ;  
 	 	 	 } ;  
 	 }  
  
 	 / / w i n d o w . m o v e T o ( s c r e e n . w i d t h + 1 0 ,   s c r e e n . h e i g h t + 1 0 ) ; 	 / / T h e   w i n d o w   h a s   a   t e n d e n c y   t o   f r e e z e ,   s o   w e   w a n t   i t   o u t   o f   t h e   w a y  
  
 	 s t a t b o x . v a l u e   =   " " ;  
 	 f o r ( j = 0 ; j < d i r s . l e n g t h ; j + + ) {  
  
 	 	 / *  
 	 	 * 	 I n i t :   o p e n   t h e   n e x t   d i r e c t o r y   a n d   g e t   o u r   r e p o r t i n g   l i s t   o f   p h o n e   n u m b e r s  
 	 	 * /  
 	 	 c d i r   =   M a k e D i r e c t o r y ( E X L , d i r s [ j ] ) ;  
 	 	 i f ( D i r R e p ) { n l i s t   =   [ ] ;  
 	 	 	 f o r ( v a r   k e y   i n   c d i r [ 1 ] ) { i f ( c d i r [ 1 ] [ k e y ]   ! =   " U "   & &   c d i r [ 1 ] [ k e y ] . s u b s t r ( 0 , 1 )   ! =   " N " ) { n l i s t . p u s h ( k e y ) ; } }  
 	 	 }  
 	 	 / / a l e r t ( " n l i s t :   " + n l i s t . j o i n ( ) ) ;  
 	 	 u p d a t e R e p o r t B l o c k s ( ) ;  
  
 	 	 / *  
 	 	 * 	 R u n   t h r o u g h   t h e   d a t a   t o   c o p y   s t u f f   i n t o   t h e   a p p r o p r i a t e   r e p o r t   p a g e s  
 	 	 * /  
 	 	 f o r ( ; i < c d a t s [ d a t i d x ] . d a t a . l e n g t h   & &   c d a t s [ d a t i d x ] . d a t a [ i ] [ 0 ] < d a t e s [ j ] ; i + + ) { }   / / g e t   t o   s t a r t i n g   d a t e  
 	 	 i f ( i = = c d a t s [ d a t i d x ] . d a t a . l e n g t h ) {   / / j u s t   i n   c a s e ,   i n p u t   s a n i t i z a t i o n  
 	 	 	 i f ( + + d a t i d x   = =   c d a t s . l e n g t h ) { b r e a k ; }  
 	 	 	 u p d a t e R e p o r t B l o c k s ( ) ;  
 	 	 }  
 	 	 / / a l e r t ( " A b o u t   t o   l o o p " ) ;  
 	 	 w h i l e ( c d a t s [ d a t i d x ] . d a t a [ i ] [ 0 ] < ( d a t e s [ j + 1 ] + 8 6 4 0 0 0 0 0 ) ) {   / / g o   t i l l   e n d i n g   d a t e ,   i n c l u s i v e  
 	 	 	 c u r r e c   =   c d a t s [ d a t i d x ] . d a t a [ i ] ;  
 	 	 	 t i m e   =   c u r r e c [ 0 ] % 8 6 4 0 0 0 0 0 ;   / / d i v i d e   o u t   d a y s ,   l e a v i n g   o n l y   t i m e   f r o m   s t a r t   o f   d a y  
 	 	 	 d u r a t i o n   =   c u r r e c [ 1 ] ;  
 	 	 	 c d u r   =   i s N a N ( d u r a t i o n ) ? - 7 2 0 0 0 0 0 : d u r a t i o n ;  
 	 	 	 o u t n u m   =   c u r r e c [ 2 ] ;  
 	 	 	 i n n u m   =   c u r r e c [ 3 ] ;  
 	 	 	 c a l l t y p e   =   " X " ;  
 	 	 	 / / a l e r t ( c d a t s [ d a t i d x ] . d a t a [ i ] [ 0 ] + "   :   " + ( d a t e s [ j + 1 ] + 8 6 4 0 0 0 0 0 ) ) ;  
 	 	 	 / / D o n ' t   j u s t   c h e c k   t h e   r a w   n u m b e r s -   c h e c k   t h e i r   s u f f i x e s   d o w n   t o   t h e   l e n g t h   o f   t h e   s h o r t e s t   k e y  
 	 	 	 / / i n   t h e   c u r r e n t   r e p o r t i n g   l i s t ,   s o   a s   t o   a v o i d   i s s u e s   w h e n   e x t r a   w e i r d   p r e f i x   d i g i t s   s h o w   u p ,   l i k e   + 3 0  
 	 	 	 f o r ( i k e y = i n n u m ; i k e y . l e n g t h & & i k e y . l e n g t h > = k e y l e n ; i k e y = i k e y . s u b s t r ( 1 ) ) {  
 	 	 	 	 / / a l e r t ( " I n c o m i n g ?   :   " + i k e y ) ;  
 	 	 	 	 i f ( N U M B E R S [ i k e y ] ) {   / / i n c o m i n g   c a l l   t o   t h i s   n u m b e r  
 	 	 	 	 	 p a r s e L i n e I n ( ) ;  
 	 	 	 	 	 b r e a k ;  
 	 	 	 	 }  
 	 	 	 }  
 	 	 	 f o r ( o k e y = o u t n u m ; o k e y . l e n g t h & & o k e y . l e n g t h > = k e y l e n ; o k e y = o k e y . s u b s t r ( 1 ) ) {  
 	 	 	 	 / / a l e r t ( " O u t g o i n g ?   :   " + o k e y ) ;  
 	 	 	 	 i f ( N U M B E R S [ o k e y ] ) {   / / o u t g o i n g   c a l l   f r o m   t h i s   n u m b e r  
 	 	 	 	 	 p a r s e L i n e O u t ( ) ;  
 	 	 	 	 	 b r e a k ;  
 	 	 	 	 }  
 	 	 	 }  
 	 	 	 i + + ;  
 	 	 	 i f ( i = = c d a t s [ d a t i d x ] . d a t a . l e n g t h ) {  
 	 	 	 	 i f ( + + d a t i d x   = =   c d a t s . l e n g t h ) { b r e a k ; }  
 	 	 	 	 u p d a t e R e p o r t B l o c k s ( ) ;  
 	 	 	 }  
 	 	 }  
 	 }  
  
 	 / * f o r ( v a r   s n a m e   i n   R E P O R T ) {  
 	 	 o u t f i l e . W r i t e L i n e ( s n a m e + " : " ) ;  
 	 	 o u t f i l e . W r i t e L i n e ( R E P O R T [ s n a m e ] . j o i n ( " \ r \ n " ) ) ;  
 	 	 o u t f i l e . W r i t e L i n e ( ) ;  
 	 } * /  
  
 	 o u t f i l e . C l o s e ( ) ;  
 	 f s o   =   n u l l ;  
  
 	 / / w i n d o w . m o v e T o ( 2 0 0 , 2 0 0 ) ;   / / p u t   i t   b a c k  
  
 	 r e t u r n   R E P O R T ;  
 }  
  
 / *   E x p o r t 2 E X L ( E X L ,   R E P O R T ,   c o l o r s e t ,   s v n a m e )  
 * 	 A r g u m e n t  
 * 	 	 E X L   ( O b j e c t ) :   A n   E x c e l   a p p l i c a t i o n   o b j e c t  
 * 	 	 R E P O R T   ( O b j e c t ) :   A   r e p o r t   s t r u c t u r e  
 * 	 	 c o l o r s e t   ( A r r a y ) :   a   l i s t   o f   c o l o r   c o d e s   f o r   t h e   v a r i o u s   c a l l   t y p e s  
 * 	 	 s v n a m e   ( S t r i n g ) :   t h e   n a m e   o f   t h e   t h e   o u p u t   f i l e  
 * 	 C l o s u r e   h i d e l a t c h ,   j ,   r s h e e t ,   s h o w b l o c k  
 * 	 V a r i a b l e   G r a p h S h e e t ,   c o l l u m n s ,   c o l n a m e s ,   c r e c ,   d t ,   d t b l c k ,   f s o ,   h c ,   i ,   k ,   k e y ,   p a t h ,   r e c t y p e ,   r g b ,   s e t H i g h l i g h t ,   s p a t h  
 * 	 E x c e p t i o n   e  
 * 	 G l o b a l   A c t i v e X O b j e c t ,   W S h e l l ,   D a t e ,   i s N a N ,   p a r s e I n t  
 * /  
  
 f u n c t i o n   E x p o r t 2 E X L ( E X L , R E P O R T , c o l o r s e t , s u m n u m , s u m t i m e , s v n a m e ) {  
 	 v a r 	 f s o   =   n e w   A c t i v e X O b j e c t ( " S c r i p t i n g . F i l e S y s t e m O b j e c t " ) ,  
 	 	 p a t h   =   f s o . G e t A b s o l u t e P a t h N a m e ( " . " ) ,  
 	 	 s p a t h   =   p a t h . s u b s t r i n g ( 0 , p a t h . l a s t I n d e x O f ( " \ \ " ) )   +   " \ \ P h o n e   R e c o r d s " ,  
 	 	 G r a p h S h e e t   =   E X L . W o r k b o o k s . O p e n ( p a t h   +   " \ \ s o u r c e f i l e s \ \ p g m . t e m p l a t e " ) ,  
 	 	 h i d e l a t c h   =   0 ,  
 	 	 s h o w b l o c k   =   1 ,  
 	 	 d t   =   n e w   D a t e ( ) ,  
 	 	 c o l l u m n s ,  
 	 	 c o l n a m e s   =   [ " A " , " B " , " C " , " D " , " E " , " F " , " G " ] ,  
 	 	 r s h e e t ,  
 	 	 d t b l c k ,  
 	 	 c r e c ,  
 	 	 i , j , k ;  
 	 	  
 	 / *   " s e t H i g h l i g h t " ( r t )  
 	 * 	 O u t e r   c o l o r s e t ,   h i d e l a t c h ,   j ,   r s h e e t ,   s h o w b l o c k  
 	 * /  
 	 f u n c t i o n   s e t H i g h l i g h t ( r t ) {  
 	 	 i f ( c o l o r s e t [ r t ]   = =   " d " ) { r e t u r n   f a l s e ; }   / / d o n ' t   b o t h e r   t o   r e c o r d   t h i s   c a l l  
 	 	 i f ( c o l o r s e t [ r t ]   <   0 ) {  
 	 	 	 i f ( c o l o r s e t [ r t ]   ! =   - 1 ) { r s h e e t . R o w s ( j ) . F o n t . C o l o r   =   - 1 * c o l o r s e t [ r t ] ; }  
 	 	 	 i f ( s h o w b l o c k ) {  
 	 	 	 	 s h o w b l o c k   =   0 ;  
 	 	 	 	 h i d e l a t c h   =   j ;  
 	 	 	 }  
 	 	 } e l s e {  
 	 	 	 i f ( ! s h o w b l o c k ) {  
 	 	 	 	 s h o w b l o c k   =   1 ;  
 	 	 	 	 r s h e e t . R o w s ( h i d e l a t c h + " : " + ( j - 1 ) ) . H i d d e n   =   t r u e ;  
 	 	 	 }  
 	 	 	 i f ( c o l o r s e t [ r t ] ) { r s h e e t . R o w s ( j ) . F o n t . C o l o r   =   c o l o r s e t [ r t ] ; }  
 	 	 }   / / h i d i n g   w h o l e   b l o c k s   a t   a   t i m e   r e q u i r e s   a n   e x t r a   c h e c k   o n   e v e r y   l i n e ,  
 	 	 r e t u r n   t r u e ;   / / b u t   s u f f i c i e n t l y   f e w e r   E x c e l   i n t e r f a c e s   t h a t   i t   s t i l l   r e s u l t s   i n   a   s p e e d   i m p r o v e m e n t  
 	 }  
  
 	 f u n c t i o n   g e t C o p y T y p e ( ) {  
 	 	 v a r   r e c t y p e , h c , r g b , r c , g c , b c ;  
 	 	 t r y {   / / w e   d o   i t   w i t h   a   t r y / c a t c h   j u s t   i n   c a s e   s o m e b o d y   s c r e w e d   u p   t h e   D i r e c t o r y .  
 	 	 	 r e c t y p e   =   c r e c [ 4 ] ;   / /   A l t h o u g h ,   w e   c o u l d   c h e c k   t h a t   w h e n   r e a d i n g   t h e   d i r e c t o r y . . . .  
 	 	 	 i f ( r e c t y p e . s u b s t r ( 0 , 1 ) = = " N " ) {   / / C a l c u l a t e   c o l o r   f o r   s p e c i a l   t r a c k i n g   n u m b e r s  
 	 	 	 	 i f ( r e c t y p e . l e n g t h > 1 ) {  
 	 	 	 	 	 i f ( r e c t y p e . i n d e x O f ( " / " ) ! = - 1 ) {   / / d e c i m a l s   s e p a r a t e d   b y   s l a s h e s  
 	 	 	 	 	 	 r g b   =   r e c t y p e . s u b s t r ( 1 ) . s p l i t ( " / " ) ;  
 	 	 	 	 	 	 r c   =   p a r s e I n t ( r g b [ 0 ] , 1 0 ) ;  
 	 	 	 	 	 	 g c   =   p a r s e I n t ( r g b [ 1 ] , 1 0 ) ;  
 	 	 	 	 	 	 b c   =   p a r s e I n t ( r g b [ 2 ] , 1 0 ) ;  
 	 	 	 	 	 } e l s e { 	 	 	 / / h e x   v a l u e  
 	 	 	 	 	 	 r c   =   p a r s e I n t ( r e c t y p e . s u b s t r ( 1 , 2 ) , 1 6 ) ;  
 	 	 	 	 	 	 g c   =   p a r s e I n t ( r e c t y p e . s u b s t r ( 3 , 2 ) , 1 6 ) ;  
 	 	 	 	 	 	 b c   =   p a r s e I n t ( r e c t y p e . s u b s t r ( 5 , 2 ) , 1 6 ) ;  
 	 	 	 	 	 }  
 	 	 	 	 	 i f ( r c   >   2 5 5 ) { r c   =   2 5 5 ; }  
 	 	 	 	 	 i f ( g c   >   2 5 5 ) { g c   =   2 5 5 ; }  
 	 	 	 	 	 i f ( b c   >   2 5 5 ) { b c   =   2 5 5 ; }  
 	 	 	 	 	 h c   =   r c + 2 5 6 * g c + 6 5 5 3 6 * b c ;  
 	 	 	 	 	 i f ( h c ) { r s h e e t . R o w s ( j ) . F o n t . C o l o r   =   h c ; }  
 	 	 	 	 } e l s e { r e t u r n   s e t H i g h l i g h t ( 3 2 ) ; }  
 	 	 	 } e l s e { r e t u r n   s e t H i g h l i g h t ( p a r s e I n t ( r e c t y p e , 1 0 ) ) ; }  
 	 	 } c a t c h ( e ) { r e t u r n   s e t H i g h l i g h t ( r e c t y p e . s u b s t r ( 0 , 1 ) = = " N " ? 3 2 : p a r s e I n t ( r e c t y p e , 1 0 ) ) ; }  
 	 	 r e t u r n   t r u e ;  
 	 }  
  
 	 i f ( ! f s o . F o l d e r E x i s t s ( s p a t h ) ) { f s o . C r e a t e F o l d e r ( s p a t h ) ; }  
 	 s p a t h   + =   " \ \ " ;  
 	 i f ( s v n a m e = = = " " ) { s v n a m e = " P h R e p " + d t . g e t F u l l Y e a r ( ) + " _ " + ( d t . g e t M o n t h ( ) + 1 ) + " _ " + d t . g e t D a t e ( ) + " . x l s x " ; }  
 	 G r a p h S h e e t . S a v e A s ( s p a t h   +   s v n a m e ) ;  
 	 / / w i n d o w . m o v e T o ( s c r e e n . w i d t h + 1 0 ,   s c r e e n . h e i g h t + 1 0 ) ; 	 / / T h e   w i n d o w   h a s   a   t e n d e n c y   t o   f r e e z e ,   s o   w e   w a n t   i t   o u t   o f   t h e   w a y  
 	  
 	 c o l l u m n s   =   [ 	 n e w   A c t i v e X O b j e c t ( " S c r i p t i n g . D i c t i o n a r y " ) ,  
 	 	 	 n e w   A c t i v e X O b j e c t ( " S c r i p t i n g . D i c t i o n a r y " ) ,  
 	 	 	 n e w   A c t i v e X O b j e c t ( " S c r i p t i n g . D i c t i o n a r y " ) ,  
 	 	 	 n e w   A c t i v e X O b j e c t ( " S c r i p t i n g . D i c t i o n a r y " ) ,  
 	 	 	 n e w   A c t i v e X O b j e c t ( " S c r i p t i n g . D i c t i o n a r y " ) ,  
 	 	 	 n e w   A c t i v e X O b j e c t ( " S c r i p t i n g . D i c t i o n a r y " ) ,  
 	 	 	 n e w   A c t i v e X O b j e c t ( " S c r i p t i n g . D i c t i o n a r y " )  
 	 	 ] ;  
 	 	  
 	 f o r ( v a r   k e y   i n   R E P O R T ) { i f ( R E P O R T . h a s O w n P r o p e r t y ( k e y ) ) {  
 	 	 v a r   b l o c k   =   R E P O R T [ k e y ] ;  
 	 	 i f ( b l o c k . l i s t . l e n g t h ) {  
 	 	 	 G r a p h S h e e t . S h e e t s ( " 1 " ) . C o p y ( n u l l , G r a p h S h e e t . W o r k s h e e t s ( G r a p h S h e e t . W o r k s h e e t s . C o u n t ) ) ;  
 	 	 	 d t b l c k   =   b l o c k . l i s t ;  
 	 	 	 r s h e e t   =   G r a p h S h e e t . A c t i v e S h e e t ;  
 	 	 	 r s h e e t . n a m e   =   k e y ;  
 	 	 	 j = 2 ;  
 	 	 	 f o r ( i = 5 ; i - - ; ) { c o l l u m n s [ i ] . r e m o v e A l l ( ) ; }  
 	 	 	 f o r ( i = 0 , k = d t b l c k . l e n g t h ; k - - ; i + + ) {  
 	 	 	 	 c r e c   =   d t b l c k [ i ] ;  
 	 	 	 	 i f ( g e t C o p y T y p e ( ) ) {  
 	 	 	 	 	 d t . s e t T i m e ( c r e c [ 0 ] ) ;  
 	 	 	 	 	 c o l l u m n s [ 0 ] . a d d ( i , c r e c [ 2 ] ) ;  
 	 	 	 	 	 c o l l u m n s [ 1 ] . a d d ( i , c r e c [ 3 ] ) ;  
 	 	 	 	 	 c o l l u m n s [ 2 ] . a d d ( i , [ " J a n   " , " F e b   " , " M a r   " , " A p r   " , " M a y   " , " J u n   " , " J u l   " , " A u g   " , " S e p t   " , " O c t   " , " N o v   " , " D e c   " ] [ d t . g e t M o n t h ( ) ] + d t . g e t D a t e ( ) + " ,   " + d t . g e t F u l l Y e a r ( ) ) ;  
 	 	 	 	 	 c o l l u m n s [ 3 ] . a d d ( i , d t . g e t H o u r s ( ) + " : " + d t . g e t M i n u t e s ( ) + " : " + d t . g e t S e c o n d s ( ) ) ;  
 	 	 	 	 	 i f ( i s N a N ( c r e c [ 1 ] ) ) { c o l l u m n s [ 4 ] . a d d ( i , " M s g " ) ; }  
 	 	 	 	 	 e l s e {  
 	 	 	 	 	 	 d t . s e t T i m e ( c r e c [ 1 ] ) ;  
 	 	 	 	 	 	 c o l l u m n s [ 4 ] . a d d ( i , d t . g e t H o u r s ( ) + " : " + d t . g e t M i n u t e s ( ) + " : " + d t . g e t S e c o n d s ( ) ) ;  
 	 	 	 	 	 }  
 	 	 	 	 	 j + + ;  
 	 	 	 	 }  
 	 	 	 }  
 	 	 	 j - - ;  
 	 	 	 i f ( j < 2 ) { c o n t i n u e ; }   / / t h i s   s h e e t   i s   e m p t y  
 	 	 	 i f ( ! s h o w b l o c k ) {  
 	 	 	 	 s h o w b l o c k   =   1 ;  
 	 	 	 	 r s h e e t . R o w s ( h i d e l a t c h + " : " + j ) . H i d d e n   =   t r u e ;  
 	 	 	 }  
  
 	 	 	 f o r ( i = 5 ; i - - ; ) { r s h e e t . R a n g e ( c o l n a m e s [ i ] + " 2 : " + c o l n a m e s [ i ] + j ) . V a l u e   =   E X L . W o r k s h e e t F u n c t i o n . T r a n s p o s e ( c o l l u m n s [ i ] . I t e m s ( ) ) ; }  
 	 	 	 W S h e l l . P o p u p ( " C o m p l e t e d   p a g e   f o r   " + k e y , 1 , " E x p o r t e r - > E x c e l " , 6 4 ) ;  
 	 	 }  
 	 	 v a r   s u m L i s t   =   [ ] ;  
 	 	 d t b l c k   =   b l o c k . s u m m a r y ;  
 	 	 f o r ( v a r   p h n a m e   i n   d t b l c k ) { i f ( d t b l c k . h a s O w n P r o p e r t y ( p h n a m e ) ) {  
 	 	 	 c r e c   =   d t b l c k [ p h n a m e ] ;  
 	 	 	 v a r   c c o u n t   =   c r e c [ 0 ] + c r e c [ 2 ] ;  
 	 	 	 v a r   c d u r   =   c r e c [ 1 ] + c r e c [ 3 ] + 7 2 0 0 0 0 0 ;  
 	 	 	 i f ( ( c c o u n t   > =   s u m n u m )   | |   ( c d u r   > =   s u m t i m e ) ) { s u m L i s t . p u s h ( [ p h n a m e , c c o u n t , c d u r ] . c o n c a t ( c r e c ) ) ; }  
 	 	 } }  
 	 	 i f ( s u m L i s t . l e n g t h ) {  
 	 	 	 s u m L i s t . s o r t ( ) ;  
 	 	 	 G r a p h S h e e t . S h e e t s . A d d ( n u l l , G r a p h S h e e t . W o r k s h e e t s ( G r a p h S h e e t . W o r k s h e e t s . C o u n t ) ) ;  
 	 	 	 r s h e e t   =   G r a p h S h e e t . A c t i v e S h e e t ;  
 	 	 	 i = 1 ;  
 	 	 	 j = 2 ;  
 	 	 	 p h n a m e   =   " S u m .   " + k e y . s u b s t r ( 0 , 2 6 ) ;  
 	 	 	 d o {  
 	 	 	 	 t r y {  
 	 	 	 	 	 r s h e e t . n a m e   =   p h n a m e ;  
 	 	 	 	 	 i = 0 ;  
 	 	 	 	 } c a t c h ( e ) {  
 	 	 	 	 	 p h n a m e   =   p h n a m e . s u b s t r ( 0 , 3 1 - M a t h . c e i l ( M a t h . l o g ( j ) / M a t h . l o g ( 1 0 ) ) ) + j ;  
 	 	 	 	 	 j + + ;  
 	 	 	 	 }  
 	 	 	 } w h i l e ( i ) ;  
 	 	 	 r s h e e t . C e l l s . N u m b e r F o r m a t   =   " @ " ;  
 	 	 	 f o r ( i = 7 ; i - - ; ) { c o l l u m n s [ i ] . r e m o v e A l l ( ) ; }  
 	 	 	 c o l l u m n s [ 0 ] . a d d ( - 1 , " P h o n e   N u m b e r " ) ;  
 	 	 	 c o l l u m n s [ 1 ] . a d d ( - 1 , " T o t a l   C a l l s " ) ;  
 	 	 	 c o l l u m n s [ 2 ] . a d d ( - 1 , " T o t a l   T i m e " ) ;  
 	 	 	 c o l l u m n s [ 3 ] . a d d ( - 1 , " I n c o m i n g   C a l l s " ) ;  
 	 	 	 c o l l u m n s [ 4 ] . a d d ( - 1 , " I n c o m i n g   T i m e " ) ;  
 	 	 	 c o l l u m n s [ 5 ] . a d d ( - 1 , " O u t g o i n g   C a l l s " ) ;  
 	 	 	 c o l l u m n s [ 6 ] . a d d ( - 1 , " O u t g o i n g   T i m e " ) ;  
 	 	 	 f o r ( i = 0 , k = s u m L i s t . l e n g t h ; i < k ; i + + ) {  
 	 	 	 	 c r e c   =   s u m L i s t [ i ] ;  
 	 	 	 	 c o l l u m n s [ 0 ] . a d d ( i , c r e c [ 0 ] ) ;  
 	 	 	 	 c o l l u m n s [ 1 ] . a d d ( i , c r e c [ 1 ] ) ;  
 	 	 	 	 d t . s e t T i m e ( c r e c [ 2 ] ) ;  
 	 	 	 	 c o l l u m n s [ 2 ] . a d d ( i , d t . g e t H o u r s ( ) + " : " + d t . g e t M i n u t e s ( ) + " : " + d t . g e t S e c o n d s ( ) ) ;  
 	 	 	 	 c o l l u m n s [ 3 ] . a d d ( i , c r e c [ 3 ] ) ;  
 	 	 	 	 d t . s e t T i m e ( c r e c [ 4 ] ) ;  
 	 	 	 	 c o l l u m n s [ 4 ] . a d d ( i , d t . g e t H o u r s ( ) + " : " + d t . g e t M i n u t e s ( ) + " : " + d t . g e t S e c o n d s ( ) ) ;  
 	 	 	 	 c o l l u m n s [ 5 ] . a d d ( i , c r e c [ 5 ] ) ;  
 	 	 	 	 d t . s e t T i m e ( c r e c [ 6 ] ) ;  
 	 	 	 	 c o l l u m n s [ 6 ] . a d d ( i , d t . g e t H o u r s ( ) + " : " + d t . g e t M i n u t e s ( ) + " : " + d t . g e t S e c o n d s ( ) ) ;  
 	 	 	 }  
 	 	 	 f o r ( i = 7 , k + + ; i - - ; ) {  
 	 	 	 	 r s h e e t . R a n g e ( c o l n a m e s [ i ] + " 1 : " + c o l n a m e s [ i ] + k ) . V a l u e   =   E X L . W o r k s h e e t F u n c t i o n . T r a n s p o s e ( c o l l u m n s [ i ] . I t e m s ( ) ) ;  
 	 	 	 }  
 	 	 	 r s h e e t . C o l u m n s . A u t o F i t ;  
 	 	 }  
 	 } }  
 	 G r a p h S h e e t . S a v e ( ) ;  
 	 G r a p h S h e e t . C l o s e ( f a l s e ) ;  
 	 f s o   =   n u l l ;  
 	 / / w i n d o w . m o v e T o ( 2 0 0 , 2 0 0 ) ;  
 }  
  
 / *   E x p o r t 2 T e x t ( R E P O R T ,   s v n a m e )  
 * 	 A r g u m e n t  
 * 	 	 R E P O R T   ( O b j e c t ) :   A   r e p o r t   s t r u c t u r e  
 * 	 	 s v n a m e   ( S t r i n g ) :   t h e   n a m e   o f   t h e   t h e   o u p u t   f i l e  
 * 	 G l o b a l   A c t i v e X O b j e c t ,   W S h e l l ,   D a t e ,   i s N a N ,   p a r s e I n t  
 * /  
  
 f u n c t i o n   E x p o r t 2 T e x t ( R E P O R T , s u m n u m , s u m t i m e , s v n a m e ) {  
 	 v a r 	 f s o   =   n e w   A c t i v e X O b j e c t ( " S c r i p t i n g . F i l e S y s t e m O b j e c t " ) ,  
 	 	 p a t h   =   f s o . G e t A b s o l u t e P a t h N a m e ( " . " ) ,  
 	 	 s p a t h   =   p a t h . s u b s t r i n g ( 0 , p a t h . l a s t I n d e x O f ( " \ \ " ) )   +   " \ \ P h o n e   R e c o r d s " ,  
 	 	 t e x t f i l e ,  
 	 	 d t   =   n e w   D a t e ( ) ,  
 	 	 c r e c ;  
 	 i f ( ! f s o . F o l d e r E x i s t s ( s p a t h ) ) { f s o . C r e a t e F o l d e r ( s p a t h ) ; }  
 	 s p a t h   + =   " \ \ " ;  
 	 i f ( ! s v n a m e ) { s v n a m e = " P h R e p " + d t . g e t F u l l Y e a r ( ) + " _ " + ( d t . g e t M o n t h ( ) + 1 ) + " _ " + d t . g e t D a t e ( ) ; }  
 	 t e x t f i l e   =   f s o . O p e n T e x t F i l e ( s p a t h + " \ \ " + s v n a m e + " . t x t " ,   2 ,   t r u e ,   - 1 ) ;   / / o p e n   f o r   w r i t i n g ,   c r e a t e   n e w   f i l e ,   u n i c o d e   f o r m a t  
 	 f o r ( v a r   k e y   i n   R E P O R T ) { i f ( R E P O R T . h a s O w n P r o p e r t y ( k e y ) ) {  
 	 	 v a r   i , b l o c k   =   R E P O R T [ k e y ] ;  
 	 	 i f ( b l o c k . l i s t . l e n g t h = = = 0   & &   b l o c k . s u m m a r y . l e n g t h = = = 0 ) { c o n t i n u e ; }  
 	 	 t e x t f i l e . W r i t e L i n e ( k e y + " : " ) ;  
 	 	 i f ( b l o c k . l i s t . l e n g t h ) { t e x t f i l e . W r i t e L i n e ( " \ t C a l l   L i s t : " ) ;  
 	 	 	 f o r ( i = 0 ; i < b l o c k . l i s t . l e n g t h ; i + + ) {  
 	 	 	 	 c r e c   =   b l o c k . l i s t [ i ]  
 	 	 	 	 d t . s e t T i m e ( c r e c [ 0 ] ) ;  
 	 	 	 	 t e x t f i l e . W r i t e ( " \ t " + [ " J a n   " , " F e b   " , " M a r   " , " A p r   " , " M a y   " , " J u n   " , " J u l   " , " A u g   " , " S e p t   " , " O c t   " , " N o v   " , " D e c   " ] [ d t . g e t M o n t h ( ) ] + d t . g e t D a t e ( ) + " ,   " + d t . g e t F u l l Y e a r ( ) ) ;  
 	 	 	 	 t e x t f i l e . W r i t e ( " \ t " + c r e c [ 2 ] + " \ t " + c r e c [ 3 ] ) ;  
 	 	 	 	 t e x t f i l e . W r i t e ( " \ t " + d t . g e t H o u r s ( ) + " : " + d t . g e t M i n u t e s ( ) + " : " + d t . g e t S e c o n d s ( ) ) ;  
 	 	 	 	 i f ( i s N a N ( c r e c [ 1 ] ) ) { t e x t f i l e . W r i t e L i n e ( " \ t M s g " ) ; }  
 	 	 	 	 e l s e {  
 	 	 	 	 	 d t . s e t T i m e ( c r e c [ 1 ] ) ;  
 	 	 	 	 	 t e x t f i l e . W r i t e L i n e ( " \ t " + d t . g e t H o u r s ( ) + " : " + d t . g e t M i n u t e s ( ) + " : " + d t . g e t S e c o n d s ( ) ) ;  
 	 	 	 	 }  
 	 	 	 }  
 	 	 	 W S h e l l . P o p u p ( " C o m p l e t e d   p a g e   f o r   " + k e y , 1 , " E x p o r t e r - > T e x t " , 6 4 ) ;  
 	 	 }  
 	 	 v a r   s u m L i s t   =   [ ] ;  
 	 	 v a r   d t b l c k   =   b l o c k . s u m m a r y ;  
 	 	 f o r ( v a r   p h n a m e   i n   d t b l c k ) { i f ( d t b l c k . h a s O w n P r o p e r t y ( p h n a m e ) ) {  
 	 	 	 c r e c   =   d t b l c k [ p h n a m e ] ;  
 	 	 	 v a r   c c o u n t   =   c r e c [ 0 ] + c r e c [ 2 ] ;  
 	 	 	 v a r   c d u r   =   c r e c [ 1 ] + c r e c [ 3 ] + 7 2 0 0 0 0 0 ;  
 	 	 	 i f ( ( c c o u n t   > =   s u m n u m )   | |   ( c d u r   > =   s u m t i m e ) ) { s u m L i s t . p u s h ( [ p h n a m e , c c o u n t , c d u r ] . c o n c a t ( c r e c ) ) ; }  
 	 	 } }  
 	 	 i f ( s u m L i s t . l e n g t h ) {  
 	 	 	 s u m L i s t . s o r t ( ) ;  
 	 	 	 t e x t f i l e . W r i t e L i n e ( " \ t C a l l   S u m m a r y : " ) ;  
 	 	 	 t e x t f i l e . W r i t e L i n e ( " \ t N u m b e r : \ t \ t T o t a l   C a l l s : \ t T o t a l   T i m e : \ t I n   C a l l s : \ t I n   T i m e : \ t O u t   C a l l s : \ t O u t   T i m e : " ) ;  
 	 	 	 f o r ( i = 0 ; i < s u m L i s t . l e n g t h ; i + + ) {  
 	 	 	 	 c r e c   =   s u m L i s t [ i ] ;  
 	 	 	 	 t e x t f i l e . W r i t e ( " \ t " + c r e c [ 0 ] ) ;  
 	 	 	 	 t e x t f i l e . W r i t e ( " \ t \ t " + c r e c [ 1 ] ) ;  
 	 	 	 	 d t . s e t T i m e ( c r e c [ 2 ] ) ;  
 	 	 	 	 t e x t f i l e . W r i t e ( " \ t " + d t . g e t H o u r s ( ) + " : " + d t . g e t M i n u t e s ( ) + " : " + d t . g e t S e c o n d s ( ) ) ;  
 	 	 	 	 t e x t f i l e . W r i t e ( " \ t \ t " + c r e c [ 3 ] ) ;  
 	 	 	 	 d t . s e t T i m e ( c r e c [ 4 ] ) ;  
 	 	 	 	 t e x t f i l e . W r i t e ( " \ t \ t " + d t . g e t H o u r s ( ) + " : " + d t . g e t M i n u t e s ( ) + " : " + d t . g e t S e c o n d s ( ) ) ;  
 	 	 	 	 t e x t f i l e . W r i t e ( " \ t \ t " + c r e c [ 5 ] ) ;  
 	 	 	 	 d t . s e t T i m e ( c r e c [ 6 ] ) ;  
 	 	 	 	 t e x t f i l e . W r i t e L i n e ( " \ t \ t " + d t . g e t H o u r s ( ) + " : " + d t . g e t M i n u t e s ( ) + " : " + d t . g e t S e c o n d s ( ) ) ;  
 	 	 	 }  
 	 	 }  
 	 	 t e x t f i l e . W r i t e L i n e ( ) ;  
 	 } }  
 	 t e x t f i l e . C l o s e ( ) ;  
 	 f s o = n u l l ;  
 	 / / w i n d o w . m o v e T o ( 2 0 0 , 2 0 0 ) ;  
 }  
 