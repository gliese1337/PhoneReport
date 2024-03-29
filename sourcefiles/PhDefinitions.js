/ * P h D e f i n i t i o n s . j s  
 * 	 C o n t a i n s   g l o b a l   v a r i a b l e   d e f i n i t i o n s .  
 * /  
  
 v a r 	 E x c e l ,  
 	 C D A T A ,  
 	 s e t w i n , r e p w i n ,  
 	 r e p D i r s   =   [ ] ,  
 	 r e p D a t e s   =   [ ] ,  
 	 r e p N u m s   =   [ ] ,  
 	 r e p N u m s T e x t ,  
 	 r e p O u t T y p e ,  
 	 s u m M i n T i m e s ,  
 	 s u m M i n D u r ,  
 	 W S h e l l ,  
 	 p a r s e r S c r i p t ,  
 	 e x p f i l e n a m e = " " ,  
 	 p s r f i l e s ,  
 	 f c o n f ,  
 	 s e t n a m e   =   " D e f a u l t " ;  
  
 f u n c t i o n   g e t T y p e d F i l e s ( s r c f o l d e r , t y p e f i l t e r ) {  
 	 v a r   f c , c f p a t h , c f n a m e , f l i s t = [ ] ;  
 	 f o r ( v a r   f c   =   n e w   E n u m e r a t o r ( s r c f o l d e r . f i l e s ) ; ! f c . a t E n d ( ) ; f c . m o v e N e x t ( ) ) {  
 	 	 c f p a t h   =   f c . i t e m ( ) + " " ;  
 	 	 c f n a m e   =   c f p a t h . s u b s t r ( c f p a t h . l a s t I n d e x O f ( " \ \ " ) + 1 ) ;  
 	 	 i f ( t y p e f i l t e r ( c f n a m e ) )   f l i s t . p u s h ( [ c f n a m e . s u b s t r ( 0 , c f n a m e . l a s t I n d e x O f ( " . " ) ) , c f p a t h ] ) ;  
 	 }  
 	 r e t u r n   f l i s t ;  
 }  
  
 f u n c t i o n   g e t P a r s e r F i l e s ( f s o ) {  
 	 v a r   p a t h   =   f s o . G e t A b s o l u t e P a t h N a m e ( " . " ) + " \ \ s o u r c e f i l e s " ;  
 	 p s r f i l e s   =   g e t T y p e d F i l e s ( f s o . G e t F o l d e r ( p a t h ) , f u n c t i o n ( f n a m e ) {  
 	 	 v a r   f t y p e   =   f n a m e . s p l i t ( " . " ) ;  
 	 	 r e t u r n   ( f t y p e [ f t y p e . l e n g t h - 1 ] = = " j s "   & &   f t y p e [ 0 ] . s u b s t r ( 0 , 5 ) = = " P H I I _ " ) ;  
 	 } ) ;  
 	 p s r f i l e s . s o r t ( ) ;  
 }  
  
 v a r   t i m e L B o u n d   =   D a t e . p a r s e ( " 1 9 7 0 / 0 1 / 0 1   6 : 0 0 : 0 0 " ) ;   / / f o r   c o m p a r i s o n   w i t h   c a l l   t i m e s  
 v a r   t i m e U B o u n d   =   D a t e . p a r s e ( " 1 9 7 0 / 0 1 / 0 1   2 2 : 1 0 : 0 0 " ) ;  
 v a r   d u r B o u n d   =   D a t e . p a r s e ( " 1 9 7 0 / 0 1 / 0 1   0 : 0 7 : 3 0 " ) ;  
  
 / * 	 N .   S p e c i a l l y   t r a c k e d   n u m b e r  
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
 v a r   c o l o r   =   [  
 	 - 1 , 	 	 / /   0 	 N o t   M i s s i o n a r y ,   u n d e r t i m e 	 h i d d e n  
 	 0 , 	 	 / /   1 	 N o t   M i s s i o n a r y ,   o v e r t i m e 	 B L A C K  
 	 0 , 	 	 / /   2   N o t   M i s s i o n a r y ,   o v e r t i m e 	 B L A C K  
 	 0 , 	 	 / /   3   N o t   M i s s i o n a r y ,   o v e r t i m e 	 B L A C K  
 	 " d " , 	 	 / /   4   A l l o w e d ,   u n d e r t i m e 	 	 d o n ' t   c o p y  
 	 1 6 7 1 1 6 8 0 , 	 / /   5   A l l o w e d ,   o v e r t i m e 	 	 B L U E  
 	 1 6 7 1 1 6 8 0 , 	 / /   6   A l l o w e d ,   o v e r t i m e 	 	 B L U E  
 	 1 6 7 1 1 6 8 0 , 	 / /   7   A l l o w e d ,   o v e r t i m e 	 	 B L U E  
 	 " d " , 	 	 / /   8   A l l o w e d ,   u n d e r t i m e 	 	 d o n ' t   c o p y  
 	 1 6 7 1 1 6 8 0 , 	 / /   9   A l l o w e d ,   o v e r t i m e 	 	 B L U E  
 	 1 6 7 1 1 6 8 0 , 	 / /   1 0   A l l o w e d ,   o v e r t i m e 	 B L U E  
 	 1 6 7 1 1 6 8 0 , 	 / /   1 1   A l l o w e d ,   o v e r t i m e 	 B L U E  
 	 " d " , 	 	 / /   1 2   A l l o w e d ,   u n d e r t i m e 	 d o n ' t   c o p y  
 	 1 6 7 1 1 6 8 0 , 	 / /   1 3   A l l o w e d ,   o v e r t i m e 	 B L U E  
 	 1 6 7 1 1 6 8 0 , 	 / /   1 4   A l l o w e d ,   o v e r t i m e 	 B L U E  
 	 1 6 7 1 1 6 8 0 , 	 / /   1 5   A l l o w e d ,   o v e r t i m e 	 B L U E  
 	 " d " , 	 	 / /   1 6   A l l o w e d ,   u n d e r t i m e 	 d o n ' t   c o p y  
 	 1 6 7 1 1 6 8 0 , 	 / /   1 7   A l l o w e d ,   o v e r t i m e 	 B L U E  
 	 1 6 7 1 1 6 8 0 , 	 / /   1 8   A l l o w e d ,   o v e r t i m e 	 B L U E  
 	 1 6 7 1 1 6 8 0 , 	 / /   1 9   A l l o w e d ,   o v e r t i m e 	 B L U E  
 	 3 2 0 0 0 , 	 	 / /   2 0   D i s a l l o w e d ,   u n d e r t i m e 	 D A R K   G R E E N  
 	 6 5 2 8 0 , 	 	 / /   2 1   D i s a l l o w e d ,   o v e r t i m e 	 L I G H T   G R E E N  
 	 6 5 2 8 0 , 	 	 / /   2 2   D i s a l l o w e d ,   o v e r t i m e 	 L I G H T   G R E E N  
 	 6 5 2 8 0 , 	 	 / /   2 3   D i s a l l o w e d ,   o v e r t i m e 	 L I G H T   G R E E N  
 	 3 2 0 0 0 , 	 	 / /   2 4   D i s a l l o w e d ,   u n d e r t i m e 	 D A R K   G R E E N  
 	 6 5 2 8 0 , 	 	 / /   2 5   D i s a l l o w e d ,   o v e r t i m e 	 L I G H T   G R E E N  
 	 6 5 2 8 0 , 	 	 / /   2 6   D i s a l l o w e d ,   o v e r t i m e 	 L I G H T   G R E E N  
 	 6 5 2 8 0 , 	 	 / /   2 7   D i s a l l o w e d ,   o v e r t i m e 	 L I G H T   G R E E N  
 	 " d " , 	 	 / /   2 8   A l l o w e d ,   u n d e r t i m e 	 d o n ' t   c o p y  
 	 1 6 7 1 1 6 8 0 , 	 / /   2 9   A l l o w e d ,   o v e r t i m e 	 B L U E  
 	 1 6 7 1 1 6 8 0 , 	 / /   3 0   A l l o w e d ,   o v e r t i m e 	 B L U E  
 	 1 6 7 1 1 6 8 0 , 	 / /   3 1   A l l o w e d ,   o v e r t i m e 	 B L U E  
 	 2 5 5 , 	 	 / /   3 2   S p e c i a l   T r a c k i n g 	 	 R E D  
 	 ] ;  
  
 f u n c t i o n   g e t S e t t i n g s F i l e s ( f s o ) {  
 	 v a r   p a t h   =   f s o . G e t A b s o l u t e P a t h N a m e ( " . " ) + " \ \ s o u r c e f i l e s " ;  
 	 f c o n f   =   g e t T y p e d F i l e s ( f s o . G e t F o l d e r ( p a t h ) , f u n c t i o n ( f n a m e ) {  
 	 	 v a r   f t y p e   =   f n a m e . s p l i t ( " . " ) ;  
 	 	 r e t u r n   ( f t y p e [ f t y p e . l e n g t h - 1 ] = = " c o n f i g " ) ;  
 	 } ) ;  
 	 f c o n f . s o r t ( ) ;  
 }  
  
 f u n c t i o n   a l t S e t t i n g s ( f s o , f n a m e ) {  
 	 v a r   s f   =   f s o . O p e n T e x t F i l e ( f n a m e ,   1 ,   f a l s e ,   - 1 ) ;   / / o p e n   f o r   r e a d i n g  
 	 t i m e L B o u n d   =   D a t e . p a r s e ( " 1 9 7 0 / 0 1 / 0 1   "   +   s f . R e a d L i n e ( ) ) ;  
 	 t i m e U B o u n d   =   D a t e . p a r s e ( " 1 9 7 0 / 0 1 / 0 1   "   +   s f . R e a d L i n e ( ) ) ;  
 	 d u r B o u n d   =   D a t e . p a r s e ( " 1 9 7 0 / 0 1 / 0 1   0 0 : "   +   s f . R e a d L i n e ( ) ) ;  
 	 f o r ( v a r   i = 0 ; i < 3 2 ; i + + ) {  
 	 	 v a r   s e t t i n g   =   s f . R e a d L i n e ( ) ;  
 	 	 c o l o r [ i ]   =   i s N a N ( s e t t i n g ) ? s e t t i n g : p a r s e I n t ( s e t t i n g ) ;  
 	 }  
 	 s f . C l o s e ( ) ;  
 }  
 